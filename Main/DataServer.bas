Attribute VB_Name = "DataServer"
Public IndexChr(26) As String

Public Function RecSrchGrid(IndexX As Long, TxtSaring As String, Optional Record As Long, Optional Search As Boolean) As Boolean
Dim tmpTxtSaring As String
Dim tmpIndexChr As Long, OverAll As Boolean
Dim Rn As Integer

tmpTxtSaring = UCase(left(TxtSaring, 1))
For XY = 1 To 26
    If OverAll = True Then nForm(2).CBGrid.M0_Value(1, XY - 1) = nForm(2).CBGrid.M0_Value(1, XY - 1) + 1
    If tmpTxtSaring = IndexChr(XY) Then
        For Z = nForm(2).CBGrid.M0_Value(1, XY - 1) To _
        Val(nForm(2).CBGrid.M0_Value(1, XY - 1)) + nForm(2).CBGrid.M0_Value(2, XY - 1)
            Rn = nForm(2).CBGrid.M0_Value(1, XY - 1)
            If nForm(1).CBGrid.M0_Value(IndexX, Z) = TxtSaring Then
                RecSrchGrid = True
                Record = Z
                Exit For
            End If
        Next Z
        If RecSrchGrid = True Then Exit For
        If Search = False Then
            nForm(2).CBGrid.M0_Value(2, XY - 1) = nForm(2).CBGrid.M0_Value(2, XY - 1) + 1
            OverAll = True
        Else
            Exit Function
        End If
    End If
Next XY
If RecSrchGrid = False Then Record = Rn
End Function

Function RecRemoveIndex(TxtSaring As String, Optional ValStr As Boolean)
    TxtSaring = UCase(left(TxtSaring, 1))
    For XY = 1 To nForm(2).CBGrid.CountY
        If OverAll = True Then nForm(2).CBGrid.M0_Value(1, XY - 1) = nForm(2).CBGrid.M0_Value(1, XY - 1) - 1
        If TxtSaring = IndexChr(XY) Then
            nForm(2).CBGrid.M0_Value(2, XY - 1) = nForm(2).CBGrid.M0_Value(2, XY - 1) - 1
            OverAll = True
        End If
    Next XY
End Function
