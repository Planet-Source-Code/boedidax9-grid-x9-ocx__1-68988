VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "*\A..\..\Main\Project2.vbp"
Begin VB.Form Form1 
   Caption         =   "Grid - X9"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Pointer"
      Height          =   1695
      Index           =   3
      Left            =   4920
      TabIndex        =   27
      Top             =   8880
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton Command6 
         Caption         =   "OK"
         Height          =   375
         Left            =   3600
         TabIndex        =   36
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   7
         Left            =   4200
         TabIndex        =   34
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   6
         Left            =   3120
         TabIndex        =   32
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   5
         Left            =   1680
         TabIndex        =   29
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   480
         TabIndex        =   28
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Count X and Y:"
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Count:"
         Height          =   210
         Index           =   10
         Left            =   3120
         TabIndex        =   53
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Count: "
         Height          =   210
         Index           =   9
         Left            =   480
         TabIndex        =   52
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Py2"
         Height          =   210
         Index           =   7
         Left            =   3840
         TabIndex        =   35
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Py1"
         Height          =   210
         Index           =   6
         Left            =   2760
         TabIndex        =   33
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Px2"
         Height          =   210
         Index           =   5
         Left            =   1320
         TabIndex        =   31
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Px1"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame3"
      Height          =   1695
      Index           =   6
      Left            =   2520
      TabIndex        =   50
      Top             =   9120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton Command4 
         Caption         =   "Test Refresh"
         Height          =   495
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame3"
      Height          =   1695
      Index           =   5
      Left            =   2520
      TabIndex        =   49
      Top             =   9120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H80000007&
      Height          =   330
      Left            =   11280
      TabIndex        =   48
      Text            =   "Select"
      Top             =   8520
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame3"
      Height          =   1695
      Index           =   4
      Left            =   2520
      TabIndex        =   47
      Top             =   9120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   330
      Left            =   120
      TabIndex        =   45
      Top             =   8520
      Width           =   11115
   End
   Begin VB.Frame Frame7 
      Height          =   1575
      Left            =   11040
      TabIndex        =   39
      Top             =   5040
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton Command9 
         Caption         =   "Apply"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   44
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto (Count)"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   42
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   8
         Left            =   1680
         TabIndex        =   40
         Text            =   "300"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Range Or Count"
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Info Grid-X9 [Fixed X]"
      Height          =   1695
      Left            =   120
      TabIndex        =   23
      Top             =   8880
      Width           =   15015
      Begin VB.CommandButton Command5 
         Caption         =   "Ü"
         Height          =   210
         Left            =   11040
         TabIndex        =   26
         Top             =   1320
         Width           =   225
      End
      Begin Project2.DBGrid DBGrid2 
         Height          =   1335
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   11175
         _extentx        =   26061
         _extenty        =   873
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1335
         Left            =   11385
         TabIndex        =   25
         Top             =   240
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   2355
         _Version        =   393217
         BorderStyle     =   0
         Appearance      =   0
         TextRTF         =   $"Form1.frx":0000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Head Position Fixed X"
      Height          =   1695
      Index           =   2
      Left            =   10200
      TabIndex        =   13
      Top             =   6795
      Width           =   4935
      Begin VB.CommandButton Command8 
         Caption         =   "Option"
         Height          =   255
         Left            =   960
         TabIndex        =   38
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2280
         TabIndex        =   37
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Fixed Index"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Only Fixed Text"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   960
         TabIndex        =   15
         Text            =   "0"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fixed "
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Visible Position Fixed X"
      Height          =   1695
      Index           =   1
      Left            =   5160
      TabIndex        =   8
      Top             =   6795
      Width           =   4935
      Begin VB.CheckBox Check2 
         Caption         =   "Fixed Index"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Visible"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Only Fixed Text"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Count Visible: "
         Height          =   210
         Index           =   12
         Left            =   960
         TabIndex        =   55
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fixed "
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Over Position Fixed X"
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   6795
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   3600
         TabIndex        =   46
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Fixed Index"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   20
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Fixed Index"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Only Fixed Text"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Text            =   "6"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Text            =   "2"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fixed To 2"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fixed To 1"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grid-X9"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin Project2.DBGrid DBGrid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _extentx        =   11880
         _extenty        =   8281
      End
   End
   Begin VB.Menu Files 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu Info_Head_Fixed 
         Caption         =   "Info Head Fixed"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RTF_Info As String
Dim GridPointer_Drag As PointInvert

Private Sub Check1_Click(Index As Integer)
If Check1(3).Value = 0 Then
    Text1(8).Text = DBGrid1.FixdColnRangeEndHead_Pos
    DBGrid1.FixdColnHeadCountAuto = False
Else
    Text1(8).Text = DBGrid1.FixdColnRangeEndHead_PosAuto
    DBGrid1.FixdColnHeadCountAuto = True
End If

DBGrid1.GoDrawing
End Sub

Private Sub Combo1_Click()
Static BackIndex As Integer

MoveFarm2 Combo1.ListIndex

Frame2(BackIndex).FontBold = False
Frame2(BackIndex).ForeColor = vbBlack

Frame2(Combo1.ListIndex).FontBold = True
Frame2(Combo1.ListIndex).ForeColor = vbBlue

BackIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
Dim TMPIndex1(0 To 1) As Long
Dim iX As Integer, iXY As Integer

TMPIndex1(0) = -1
TMPIndex1(1) = -1

If Check1(0).Value = 0 And IsNumeric(Text1(0).Text) = True And IsNumeric(Text1(1).Text) = True Then  'Index
    If Check2(0).Value = 0 Then TMPIndex1(0) = Text1(0).Text Else TMPIndex1(0) = DBGrid1.FixdColnIndex_Real(Text1(0).Text)
    If Check2(1).Value = 0 Then TMPIndex1(1) = Text1(1).Text Else TMPIndex1(1) = DBGrid1.FixdColnIndex_Real(Text1(1).Text)
Else 'Text Caption
    For iXY = 0 To 1
        For iX = 0 To DBGrid1.SellCountColumn - 1
            If Text1(iXY).Text = DBGrid1.FixedColsText(iX) Then
                TMPIndex1(iXY) = iX
                Exit For
            End If
        Next iX
    Next iXY
End If

If TMPIndex1(0) > -1 And TMPIndex1(1) > -1 Then
    DBGrid1.FixdColnIndex(TMPIndex1(0)) = TMPIndex1(1)
    DBGrid1.GoDrawing
    
    If Info_Head_Fixed.Checked = False Then
        InfoFixed TMPIndex1(0), 0
        InfoFixed TMPIndex1(1), 0
    Else
        LoadDBGrid2
    End If
    InfoFixed_Info
Else
    If TMPIndex1(0) = -1 Then MsgBox "Text [Fixed To 1( " & Text1(0).Text & " )] Not Found"
    If TMPIndex1(1) = -1 Then MsgBox "Text [Fixed To 2( " & Text1(1).Text & " )] Not Found"
End If
End Sub

Sub InfoFixed(iX As Long, iY As Long)
Dim InfoText As String, nDatas As String
Dim Index_iX As Long

    If Info_Head_Fixed.Checked = False Then
        Index_iX = iX
    Else
        Index_iX = DBGrid1.FixdColnShowHead(iX)
    End If

    'DBGrid1.FixdColnVisible(TMPIndex1)
    'InfoText = "~ " & DBGrid1.FixdColnIndex(iX) & " # " & DBGrid1.FixdColnIndexShow_Count & vbCrLf
    
    InfoText = InfoText & "! " & DBGrid1.FixdColnIndex_OnReal(Index_iX) & vbCrLf
    
    If DBGrid1.FixdColnVisible(Index_iX) = True Then nDatas = "[True]" Else nDatas = ""
    InfoText = InfoText & "@ " & DBGrid1.FixdColnIndex_Real(Index_iX) & " " & nDatas & vbCrLf
    
    If DBGrid1.FixdColnIndex_ShowHead(Index_iX) > 0 Then nDatas = "[True]" Else nDatas = ""
    InfoText = InfoText & "# " & DBGrid1.FixdColnIndex_ShowHead(Index_iX) & " " & nDatas '& vbCrLf
    
    DBGrid2.FixedColsText(iX) = DBGrid1.FixedColsText(Index_iX)
    DBGrid2.SellText(iX, 0) = InfoText
    
    'DBGrid2.GoDrawing '-> Del
End Sub

Private Sub Command2_Click()
Dim TMPIndex1 As Long
Dim iX As Integer

TMPIndex1 = -1
If Check1(1).Value = 0 And IsNumeric(Text1(2).Text) = True Then 'Index
    If Check2(2).Value = 0 Then TMPIndex1 = Text1(2).Text Else TMPIndex1 = DBGrid1.FixdColnIndex(Text1(2).Text)
Else 'Text Caption
    For iX = 0 To DBGrid1.SellCountColumn - 1
        If Text1(2).Text = DBGrid1.FixedColsText(iX) Then
            TMPIndex1 = iX
            Exit For
        End If
    Next iX
End If

If TMPIndex1 > -1 Then
    DBGrid1.FixdColnVisible(TMPIndex1) = Check3.Value
    DBGrid1.GoDrawing
    
    If Info_Head_Fixed.Checked = False Then
        For iX = 0 To DBGrid1.SellCountColumn - 1
            InfoFixed Val(iX), 0
        Next iX
    Else
        LoadDBGrid2
    End If
    
    InfoFixed_Info
Else
    MsgBox "Text [Fixed( " & Text1(2).Text & " )] Not Found"
End If

Label1(12).Caption = "Count Visible: " & DBGrid1.FixdHideCountCol
End Sub

Private Sub Command3_Click()
Dim TMPIndex1 As Integer
Dim iX As Integer

TMPIndex1 = -1
If Check1(1).Value = 0 And IsNumeric(Text1(3).Text) = True Then 'Index
    If Check2(3).Value = 0 Then TMPIndex1 = Text1(3).Text Else TMPIndex1 = DBGrid1.FixdColnIndex(Text1(3).Text)
Else 'Text Caption
    For iX = 0 To DBGrid1.SellCountColumn - 1
        If Text1(3).Text = DBGrid1.FixedColsText(iX) Then
            TMPIndex1 = iX
            Exit For
        End If
    Next iX
End If

If TMPIndex1 > -1 Then
    DBGrid1.Add_FixdColnShowHead TMPIndex1
    DBGrid1.GoDrawing
    
    If Info_Head_Fixed.Checked = False Then
        For iX = 0 To DBGrid1.SellCountColumn - 1
            InfoFixed Val(iX), 0
        Next iX
    Else
        LoadDBGrid2
    End If
    
    InfoFixed_Info
Else
    MsgBox "Text [Fixed( " & Text1(2).Text & " )] Not Found"
End If
End Sub

Private Sub Command4_Click()
DBGrid1.GoDrawing
End Sub

Private Sub Command5_Click()
DBGrid1.SetFocus
PopupMenu Files
End Sub

Private Sub Command6_Click()
GridPointer_Drag.Px1 = Text1(4).Text
GridPointer_Drag.Px2 = Text1(5).Text
GridPointer_Drag.Py1 = Text1(6).Text
GridPointer_Drag.Py2 = Text1(7).Text

DBGrid1.PointerGrid_Set GridPointer_Drag
End Sub

Private Sub Command7_Click()
MsgBox "Tobe Continue"
End Sub

Private Sub Command8_Click()
If Command8.Enabled = True Then
    Frame7.Visible = True
    Command8.Enabled = False
Else
    Frame7.Visible = False
    Command8.Enabled = True
End If
End Sub

Private Sub Command9_Click(Index As Integer)
Select Case Index
Case 0
    Frame7.Visible = False
    Command8.Enabled = True
Case 1
    If Check1(3).Value = 0 Then
        DBGrid1.FixdColnRangeEndHead_Pos = Text1(8).Text
    Else
        DBGrid1.FixdColnRangeEndHead_PosAuto = Text1(8).Text
    End If
    
    DBGrid1.GoDrawing
End Select
End Sub

Private Sub DBGrid1_ChangeHAfter()
InfoFixed_Info
End Sub

Private Sub DBGrid1_MouseMoveSell(IndexX As Long, IndexY As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
InfoPointers
End Sub

Private Sub DBGrid1_MouseUpSell(IndexX As Long, IndexY As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
InfoPointers
End Sub

Sub InfoPointers()
GridPointer_Drag = DBGrid1.PointerGrid_Get
Text1(4).Text = GridPointer_Drag.Px1
Text1(5).Text = GridPointer_Drag.Px2
Text1(6).Text = GridPointer_Drag.Py1
Text1(7).Text = GridPointer_Drag.Py2

Label1(9).Caption = "Count: " & GridPointer_Drag.Px2 - GridPointer_Drag.Px1 + 1
Label1(10).Caption = "Count: " & GridPointer_Drag.Py2 - GridPointer_Drag.Py1 + 1
Label1(11).Caption = "Count X and Y: " & (GridPointer_Drag.Px2 - GridPointer_Drag.Px1 + 1) * (GridPointer_Drag.Py2 - GridPointer_Drag.Py1 + 1)
End Sub

Private Sub DBGrid1_ScrollH()
InfoFixed_Info
End Sub

Private Sub Form_Load()
Dim iXY As Long
Dim iX As Long, iY As Long

MoveFarm2 0
HScroll1.Max = Frame2.Count - 3

For iXY = 0 To Frame2.Count - 1
    Combo1.AddItem "[" & Format(iXY + 1, "0#") & "] " & Frame2(iXY).Caption
Next iXY

RichTextBox1.LoadFile App.Path & "\A.RTF"
RTF_Info = RichTextBox1.Text

iX = 0
iY = 0
            DBGrid1.Add 150, 150
            
            DBGrid1.AddList iX, iY, 1
            DBGrid1.ListType(iX, iY) = 1
            
            DBGrid1.ListRange(iX, iY, 0) = 16
                DBGrid1.AddListText iX, iY, 0, 2
                DBGrid1.ListCaption(iX, iY, 0) = "I Love VB"
                DBGrid1.ListText(iX, iY, 0, 0) = "Please"
                DBGrid1.ListText(iX, iY, 0, 1) = " Vote"
                DBGrid1.ListText(iX, iY, 0, 2) = "  Me"
        
            DBGrid1.ListRange(iX, iY, 1) = 16
                DBGrid1.AddListText iX, iY, 1, 20
                DBGrid1.ListCaption(iX, iY, 1) = "My Music"
                DBGrid1.ListText(iX, iY, 1, 0) = "Luca Turilli - New Centurys Tarantella"
                DBGrid1.ListText(iX, iY, 1, 1) = "Thy Majestie - In God We Trust"
                DBGrid1.ListText(iX, iY, 1, 2) = "Dark Moor - The Bane Of Daninsky(The Werewolf) ƒ£Ü"
                DBGrid1.ListText(iX, iY, 1, 3) = "Dark Moor - Maid Of Orleans"
                DBGrid1.ListText(iX, iY, 1, 4) = "Nightwish - Wishmaster - 11. FantasMic"
                DBGrid1.ListText(iX, iY, 1, 4) = "Nightwish - Wishmaster - 10. Dead Boy's Poem"
                DBGrid1.ListText(iX, iY, 1, 4) = "Michael Jackson Another Part Of Me"
            DBGrid1.SellHeight(0, True) = 200


iX = 0
iY = 1
            DBGrid1.AddList iX, iY, 1
            DBGrid1.ListType(iX, iY) = 1
            
            DBGrid1.ListRange(iX, iY, 0) = 16
                DBGrid1.AddListText iX, iY, 0, 0
                DBGrid1.ListCaption(iX, iY, 0) = "Comment"
                DBGrid1.ListText(iX, iY, 0, 0) = "Yeah I Like Disco"
        
            DBGrid1.ListRange(iX, iY, 1) = 16
                DBGrid1.AddListText iX, iY, 1, 0
                DBGrid1.ListCaption(iX, iY, 1) = "I Love Indonesia"
                DBGrid1.ListText(iX, iY, 1, 0) = "Bali"

            DBGrid1.SellHeight(1, True) = 200



iX = 1
iY = 0
            DBGrid1.AddList iX, iY, 0
            DBGrid1.ListType(iX, iY) = 1
            
            DBGrid1.ListRange(iX, iY, 0) = 16
                DBGrid1.AddListText iX, iY, 0, 150
                DBGrid1.ListCaption(iX, iY, 0) = "I Love VB"
                For iXY = 0 To DBGrid1.ListCountText(iX, iY, 0) - 1
                    DBGrid1.ListText(iX, iY, 0, iXY) = Format(iXY, "00#") & " BoedidaX9"
                Next iXY
Dim AddWidth As Integer
For iX = 0 To DBGrid1.SellCountColumn - 1
    DBGrid1.FixedColsText(iX) = "Fixed " & iX
    ''InfoFixed iX, 0
    
    If DBGrid1.SellWidth(iX, True) = 0 Then DBGrid1.SellWidth(iX, True) = DBGrid1.SellWidth_Def
    AddWidth = AddWidth + 1
    DBGrid1.SellWidth(iX, True) = DBGrid1.SellWidth(iX, True) + AddWidth
Next iX

DBGrid1.PictureGrid = LoadPicture(App.Path & "\Gambar\PictureIcon.bmp")
DBGrid1.SellIcon(0) = True
DBGrid1.SellIcon(3) = True
DBGrid1.SellIconIndex(3, 1) = 2

For iY = 0 To DBGrid1.SellCountRow - 1
    For iX = 0 To DBGrid1.SellCountColumn - 1
        DBGrid1.SellBackColor(iX, iY) = &HFFFFC0
        DBGrid1.SellText(iX, iY) = iX & "," & iY
        DBGrid1.SellSubBackColor(iX, iY) = &HFFFAA0
        DBGrid1.SellSubText(iX, iY) = "SUB " & iX & "," & iY
    Next iX
Next iY

DBGrid1.GoDrawing 'Master Prosesing Or Refresh

LoadDBGrid2

InfoFixed_Info

Text1(8).Text = DBGrid1.FixdColnRangeEndHead_Pos
End Sub

Sub InfoFixed_Info()
    RichTextBox1.Text = RTF_Info
    RichTextBox1.Text = RichTextBox1.Text & "" & "FixdColnIndexShow_Count [" & DBGrid1.FixdColnIndexShow_Count & "]" & vbCrLf
    RichTextBox1.Text = RichTextBox1.Text & "" & "FixdHideCountCol [" & DBGrid1.FixdHideCountCol & "]" & vbCrLf
    RichTextBox1.Text = RichTextBox1.Text & "" & "FixdColnShowHead_Count [" & DBGrid1.FixdColnShowHead_Count & "]" & vbCrLf
    RichTextBox1.Text = RichTextBox1.Text & "" & "FixdColnShowHead_RunCount [" & DBGrid1.FixdColnShowHead_RunCount & "]" & vbCrLf

    RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Please Read Read Me.DOC For Very Detail, Indonesia Language" & vbCrLf
End Sub

Private Sub HScroll1_Change()
MoveFarm2 HScroll1.Value
End Sub

Private Sub Info_Head_Fixed_Click()
    If Info_Head_Fixed.Checked = False Then
        If DBGrid1.FixdColnShowHead_Count = 0 Then
            MsgBox "Nothing Proses Because [FixdColnShowHead_Count = 0]"
            Exit Sub
        Else
            Info_Head_Fixed.Checked = True
        End If
    Else
        Info_Head_Fixed.Checked = False
    End If

    LoadDBGrid2
End Sub

Sub LoadDBGrid2()
Dim iX As Long

    DBGrid2.ClearGrid
    If Info_Head_Fixed.Checked = False Then
        DBGrid2.Add 150
    Else
        DBGrid2.Add DBGrid1.FixdColnShowHead_Count
    End If
    
    DBGrid2.SellHeight_Def = 45
    For iX = 0 To DBGrid2.SellCountColumn - 1
        If Info_Head_Fixed.Checked = False Then
            DBGrid2.FixedColsText(iX) = DBGrid1.FixedColsText(iX)
        Else
            DBGrid2.FixedColsText(iX) = DBGrid1.FixedColsText(DBGrid1.FixdColnShowHead(iX))
        End If
            
        If DBGrid2.SellWidth(iX, True) = 0 Then DBGrid2.SellWidth(iX, True) = DBGrid2.SellWidth_Def
        DBGrid2.SellWidth(iX, True) = DBGrid2.SellWidth(iX, True) + 30
        
        InfoFixed Val(iX), 0
    Next iX
    
    DBGrid2.GoDrawing
End Sub

Sub MoveFarm2(IndexS As Integer)
Dim iXY As Integer, IndexMov As Integer

If IndexS > Frame2.Count - 3 Then IndexS = Frame2.Count - 3
Frame2(IndexS).Left = Frame1.Left
Frame2(IndexS).Top = Frame1.Top + Frame1.Height + 30
Frame2(IndexS).ZOrder 0
For iXY = 1 To 2
    IndexMov = iXY + IndexS
    Frame2(IndexMov).Left = Frame2(IndexMov - 1).Left + Frame2(IndexMov - 1).Width + 100
    Frame2(IndexMov).Top = Frame2(IndexMov - 1).Top
    Frame2(IndexMov).Visible = True
    Frame2(IndexMov).ZOrder 0
Next iXY
End Sub
