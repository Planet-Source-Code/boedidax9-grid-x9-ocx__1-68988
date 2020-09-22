VERSION 5.00
Begin VB.UserControl DBGrid 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   11205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13710
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   747
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   914
   Begin VB.CommandButton Command36 
      Caption         =   "Command36"
      Height          =   495
      Left            =   10200
      TabIndex        =   58
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Command35"
      Height          =   495
      Left            =   8880
      TabIndex        =   57
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Command34"
      Height          =   495
      Left            =   9000
      TabIndex        =   56
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Command33"
      Height          =   495
      Left            =   9000
      TabIndex        =   55
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Command32"
      Height          =   495
      Left            =   7680
      TabIndex        =   54
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Command31"
      Height          =   495
      Left            =   7680
      TabIndex        =   53
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Command30"
      Height          =   495
      Left            =   480
      TabIndex        =   52
      Top             =   5880
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.HScrollBar HscHead 
      Height          =   255
      Left            =   240
      Max             =   100
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Command23"
      Height          =   1695
      Left            =   360
      TabIndex        =   44
      Top             =   1680
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "0"
      Top             =   6480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Command29"
      Height          =   855
      Left            =   3960
      TabIndex        =   50
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Command27"
      Height          =   855
      Left            =   3960
      TabIndex        =   48
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Command28"
      Height          =   855
      Left            =   3960
      TabIndex        =   49
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Command26"
      Height          =   855
      Left            =   480
      TabIndex        =   47
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Command25"
      Height          =   855
      Left            =   480
      TabIndex        =   46
      Top             =   4080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Command24"
      Height          =   855
      Left            =   480
      TabIndex        =   45
      Top             =   3240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2760
      Top             =   3840
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Command22"
      Height          =   855
      Left            =   5280
      TabIndex        =   43
      Top             =   5040
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Command21"
      Height          =   855
      Left            =   5280
      TabIndex        =   42
      Top             =   4200
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Command20"
      Height          =   495
      Left            =   1200
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   10680
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   6720
      TabIndex        =   32
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   5400
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   495
      Left            =   600
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   495
      Left            =   3000
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   495
      Left            =   4200
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Command15"
      Height          =   495
      Left            =   1800
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   495
      Left            =   8040
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   495
      Left            =   9360
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   5040
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00808080&
      Caption         =   "Command12"
      Height          =   495
      Left            =   10320
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   20
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   6360
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      ToolTipText     =   "Click here to put text"
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   5760
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   960
   End
   Begin VB.HScrollBar Hsc 
      Height          =   255
      Left            =   1920
      Max             =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7320
      Width           =   6060
   End
   Begin VB.VScrollBar Vsc 
      Height          =   5895
      Left            =   10680
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10395
      Left            =   120
      ScaleHeight     =   691
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   834
      TabIndex        =   0
      Top             =   120
      Width           =   12540
      Begin VB.PictureBox picScrollMove 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6480
         Picture         =   "UserControl1.ctx":001B
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   41
         Top             =   1320
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox picToolTipText 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   6960
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   276
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   4140
         Begin VB.CheckBox chkToolTipText 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Show"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   75
            TabIndex        =   39
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Shape shpToolTipText 
            BorderStyle     =   3  'Dot
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblToolTipText 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   38
            Top             =   360
            Width           =   630
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   4080
         Picture         =   "UserControl1.ctx":059D
         ScaleHeight     =   43
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   87
         TabIndex        =   5
         Top             =   4080
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2400
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   5160
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Command19"
         Height          =   495
         Left            =   6360
         TabIndex        =   36
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox picShow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   3000
         ScaleHeight     =   303
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   287
         TabIndex        =   35
         Top             =   3840
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Command18"
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   5880
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   24
         Top             =   8280
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4920
         Picture         =   "UserControl1.ctx":3237
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   23
         Top             =   7800
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.ListBox List3 
         Height          =   5520
         Left            =   10680
         TabIndex        =   22
         Top             =   3720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   6720
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "UserControl1.ctx":3851
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox List2 
         Height          =   5520
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ListBox List1 
         Height          =   5520
         Left            =   7920
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   12855
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   12855
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   6120
         ScaleHeight     =   7815
         ScaleWidth      =   15
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   360
         ScaleHeight     =   119
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   295
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0FF&
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   2
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   80
         X2              =   80
         Y1              =   48
         Y2              =   72
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         BorderStyle     =   2  'Dash
         BorderWidth     =   2
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFC0&
         FillStyle       =   7  'Diagonal Cross
         Height          =   30
         Left            =   1335
         Top             =   540
         Width           =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0FF&
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   2
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   88
         X2              =   88
         Y1              =   40
         Y2              =   64
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   2  'Dash
         BorderWidth     =   2
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFC0&
         Height          =   855
         Left            =   1680
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00C0FFC0&
         Height          =   825
         Left            =   3600
         Top             =   720
         Width           =   1905
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         BorderStyle     =   2  'Dash
         BorderWidth     =   2
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1455
         Top             =   660
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   135
         Left            =   0
         Picture         =   "UserControl1.ctx":3857
         Top             =   840
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   3480
         Top             =   600
         Width           =   1935
      End
      Begin VB.Image Image9 
         Height          =   300
         Left            =   5760
         Picture         =   "UserControl1.ctx":3A91
         Top             =   1800
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image Image8 
         Height          =   330
         Left            =   4440
         Picture         =   "UserControl1.ctx":3F83
         Top             =   7800
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image Image7 
         Height          =   5340
         Left            =   8280
         Picture         =   "UserControl1.ctx":459D
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1320
         Picture         =   "UserControl1.ctx":1A44F
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   5400
         Picture         =   "UserControl1.ctx":1A453
         Top             =   1800
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   135
         Index           =   0
         Left            =   5280
         Top             =   6960
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   5400
         Picture         =   "UserControl1.ctx":1A457
         Top             =   6960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   1800
         Picture         =   "UserControl1.ctx":1A45B
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   300
         Left            =   480
         Top             =   120
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Image tmpImage 
      Height          =   255
      Left            =   6240
      Top             =   10560
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "DBGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Type TRIVERTEX
    X     As Long
    Y     As Long
    R     As Integer
    G     As Integer
    b     As Integer
    Alpha As Integer
End Type

Private Type RGB
    R As Integer
    G As Integer
    b As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft  As Long
    LowerRight As Long
End Type

Private Type RECT2
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Public Type PointInvert
    Px1 As Long
    Px2 As Long
    Py1 As Long
    Py2 As Long
End Type

Private m_BackSelectedG1   As OLE_COLOR
Private m_BackSelectedG2   As OLE_COLOR


Dim CBGrid As Integer
Public AAAA As BGrid

Dim CountTpXy_X As Long, CountTpXy_Y As Long
'konvert konvert go go
'----------------------------------------------------------
'BGrid V 1.0
'By: BoedidaX9
'Email: BoedidaX9@yahoo.com
'----------------------------------------------------------
'Private Type PointInvert
'    InX1 As Single
'    InY1 As Single
'    InX2 As Single
'    InY2 As Single
'End Type
'Private InvertME() As PointInvert
Enum TypeAdd
    Fixed_X = 0
    Fixed_Y = 1
    Fixed_XY = 2
    Fixed_Sell = 3
    Fixed_Sell_X = 4
    Fixed_Sell_Y = 5
    Fixed_Sell_XY = 6
End Enum

Dim Fx As Boolean, Fy As Boolean
Dim ClickScroll As Boolean
Dim MDown As Boolean
Dim InvertDrag As Boolean

Dim DragXY As Boolean
Dim TypeDrag As String, TypeDragMove As Boolean

Dim DblX As Single, DblY As Single

Dim Pxy As PointInvert
'Dim Txy As PointInvert
Dim TPxy As PointInvert
Dim TTPxy As PointInvert

Dim HTPxy As PointInvert
'----------------------------------------------------------

Public GridLeftX As Integer
    'GridXYData(GridLeftX) to first load
'Public GridRightX As Integer
    'GridXYData(GridRightX) to end load
Public XMovGrid As Integer
    'SellWidth_Def to pic2.line and pic1.line
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
    'SellWidth_Def to pic2.line
Public YIndexGrid As Integer
    'Point + <-> Y
Public GragY As Integer, tmpDragY As Integer
    'Point [] Y to all X
Public YPointerIndex As Long, DblYPointerIndex As Long
Public SubYPointerIndex As Integer

Private SetNewGrid As NewGridXY
Private GridX() As TypeGridX '99 ----------
Private GridY() As TypeGridY '3999 --------
Private GridXYData() As TypeGridXYData  '99,3999

Private tmpPictureGrid As StdPicture
Private tmpPictureSubGridY As StdPicture


'Private GridList As GList

'Private PictureSub As Image
'Private PictureXY  As Image
Private PicErr As Boolean
Private Pusing As Boolean
'Public SetNewGrid.GridSize.GDRangeX As Integer, SetNewGrid.GridSize.GDRangeY As Integer
'Public SetNewGrid.GridSize.SellHeight_Def As Integer, SetNewGrid.GridSize.SellWidth_Def As Integer
'Public SetNewGrid.GridSize.TableHeight As Single, SetNewGrid.GridSize.TableWidth As Single

'Public GridLeftX As Integer, GridUpY As Integer
Public GridRightX As Integer, GridDownY As Integer
Public IndexGL As Integer


'Public SellSubList As New GList
'Public Tager As String
'Public SetNewGrid.GridSizePicSub.RangePicSubX1 As Integer,
'SetNewGrid.GridSizePicSub.RangePicSuby1 As Integer
'Public SetNewGrid.GridSizePicSub.RangePicSubX2 As Integer,
'SetNewGrid.GridSizePicSub.RangePicSubY2 As Integer

Dim ClikScrol As Boolean, SubClikScrol As Boolean
Dim TmpScrolY As Single, SubTmpScrolY
    
Dim ScrolY1 As Single, ScrolY2 As Single
Dim MovingY As Single, SubMovingYs As Single

Dim IndexList As Integer, CountList As Integer
Dim SubIndexList As Integer, SubCountList As Integer

Dim YListGrid As Integer
Dim SubYListGrid As Integer

Dim TMP_GLSubScrolIndex As Integer
Dim tmpPointerListHead As Integer

Public Event ClikSellIcon(IndexX As Long, IndexY As Long)
Public Event MouseDownSellIcon(IndexX As Long, IndexY As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMoveSellIcon(IndexX As Long, IndexY As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUpSellIcon(IndexX As Long, IndexY As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ClikSell(IndexX As Long, IndexY As Long)
Public Event ClikSellSub(Control As String, IndexX As Long, IndexY As Long, IndexControl As Single)

Public Event MouseUpSell(IndexX As Long, IndexY As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMoveSell(IndexX As Long, IndexY As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event KeyEnter(IndexX As Long, IndexY As Long, Text As String, Cancel As Boolean)
Public Event KeyDown(IndexX As Long, IndexY As Long, KeyCode As Integer, Shift As Integer)
'     DBGrid1_KeyDown(IndexX As Long, IndexY As Long, KeyCode As Integer, Shift As Integer)

Public Event KeyDownInput(IndexX As Long, IndexY As Long, KeyCode As Integer, Shift As Integer, InputText As String, NoHideText As Boolean)
Public Event ScrollV()
Public Event ChangeVBefore()
Public Event ChangeVAfter()
Public Event ScrollH()
Public Event ChangeHBefore()
Public Event ChangeHAfter()

Public ObjcText As TextBox

Public GridFullWidth  As Long
Public GridFullHeight As Long

Public Type TPicTollTips
    nLeft   As Single
    nWidth  As Single
    nTop    As Single
    nHeight As Single
    nText   As String
    nTitle  As String
End Type

Dim XIndexPointerText1 As Long
Dim YIndexPointerText1 As Long

Private Datass As Data
Dim J As DataObject

Dim MultiSelectXY As Integer
Dim IndexMultiSelectXY() As Boolean ', MultiSelectY As Boolean
Dim Count_IndexMultiSelect As Long
Dim On_IndexMultiSelectXY As Long
Dim Data_IndexMultiSelect As New Collection

Dim OnClickSellIncon As Boolean

Dim ShowFixCol() As Integer
Dim ConShowFixCol As Integer
Dim HideCountFixCol As Integer
Dim RealShowFixCol() As Integer
Dim HeadSellOnFixCol() As Integer ' OnHeadSellOnFixCol  'Integer
Dim ConHeadSellOnFixCol As Integer
Dim ShowOnConHeadSell As Integer

Dim HeadWidFixCol As Integer
'Dim RangeEndHead As Integer
'Private Type OnHeadSellOnFixCol
'    IndexFixCol As Integer
'    RealPosision As Integer
'End Type
Dim RangeEndHead As Integer, RangeEndHead_Pos As Integer, RangeEndHead_PosAuto As Integer, HeadCountAuto As Boolean
Dim IndexEndHead As Integer



'Dim AddHides As New Collection

Sub SaveMes()
    GetData Picture1.Tag
    Open "D:\Data.txt" For Binary As #1
    Put #1, , SetNewGrid
    Put #1, , GridX()
    Put #1, , GridY()
    Put #1, , GridXYData()
    Close #1
End Sub

Sub Haji(G As DataObject)
'Dim dbConnectionString As String, FileNames As String
'
'FileNames = App.Path & "\B13.mdb"
'dbConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & FileNames & ";" '";DefaultDir=" & DetermineDirectory(FileNames) & ";"
'ObjADO.ConnectionString = dbConnectionString
    
'set  Datass.Database.Connection.Execute dbConnectionString
End Sub
'Public Event ClikControl(Control As String, IndexX As Long, IndexY As Long)

'^^^^^^^^^^^^^^^^^^^^^^^^ Public Event OR Perubahan Pointer

'Public SellIconX1 As Integer, SellIconY1 As Integer
'Public SetNewGrid.GridSizePicSub.SellIconX2 As Integer, SetNewGrid.GridSizePicSub.SellIconY2 As Integer
'Public SetNewGrid.GridSizePicSub.SellIconPicContColms As Integer, SetNewGrid.GridSizePicSub.SellIconPicContRows As Integer
    
Function MultiSelect(Optional nSelect As Integer) As Integer
Dim ClearData_IndexMultiSelect As New Collection
   
    If nSelect < 1 Then
        MultiSelect = MultiSelectXY
    Else
        If nSelect = 1 Then
            ReDim IndexMultiSelectXY(SetNewGrid.GridXCount - 1)
            On_IndexMultiSelectXY = SetNewGrid.GridXCount - 1
        ElseIf nSelect = 2 Then
            'For Yi = Data_IndexMultiSelect.Count To 1 Step -1
            '    Data_IndexMultiSelect.Remove Yi
            'Next Yi
            Set Data_IndexMultiSelect = ClearData_IndexMultiSelect
            Count_IndexMultiSelect = 0
            ReDim IndexMultiSelectXY(SetNewGrid.GridYCount - 1)
            On_IndexMultiSelectXY = SetNewGrid.GridYCount - 1
        End If
        If nSelect < 3 Then
            MultiSelectXY = nSelect
        Else
            Count_IndexMultiSelect = 0
            ReDim IndexMultiSelectXY(Count_IndexMultiSelect)
            MultiSelectXY = 0
        End If
    End If
End Function 'TRIVERTEX

Property Get MultiSelectIndex(Indexs As Long) As Boolean
    MultiSelectIndex = IndexMultiSelectXY(Indexs)
End Property
Property Let MultiSelectIndex(Indexs As Long, ByVal NewValue As Boolean)
Dim Yi As Long

    If IndexMultiSelectXY(Indexs) = True And NewValue = False Then 'False
        Count_IndexMultiSelect = Count_IndexMultiSelect - 1
        If Indexs = On_IndexMultiSelectXY Then
            On_IndexMultiSelectXY = SetNewGrid.GridYCount - 1
            For Yi = Indexs + 1 To SetNewGrid.GridYCount - 1
                If IndexMultiSelectXY(Yi) = True Then
                    On_IndexMultiSelectXY = Yi
                    Exit For
                End If
            Next Yi
        End If
        'MsgBox Data_IndexMultiSelect.Count
        For Yi = 1 To Data_IndexMultiSelect.Count
            If Data_IndexMultiSelect.Item(Yi) = Indexs Then
                Data_IndexMultiSelect.Remove Yi
                Exit For
            End If
        Next Yi
        'GridY(Indexs).Tmp_MultiSelect = 0
    ElseIf IndexMultiSelectXY(Indexs) = False And NewValue = True Then 'True
        Count_IndexMultiSelect = Count_IndexMultiSelect + 1
        If Indexs < On_IndexMultiSelectXY Then On_IndexMultiSelectXY = Indexs
        Data_IndexMultiSelect.Add Indexs
        MsgBox Data_IndexMultiSelect.Count '& " " & ClearData_IndexMultiSelect.Count
        'GridY(Indexs).Tmp_MultiSelect = Data_IndexMultiSelect.Count
    End If
    
    IndexMultiSelectXY(Indexs) = NewValue
End Property

Property Get Enabled() As Boolean
    Enabled = Picture1.Enabled
End Property
Property Let Enabled(ByVal NewValue As Boolean)
    Picture1.Enabled = NewValue
End Property

Property Get MultiSelectIndex_On() As Long
    MultiSelectIndex_On = On_IndexMultiSelectXY
End Property
Property Let MultiSelectIndex_On(ByVal NewValue As Long)
    On_IndexMultiSelectXY = NewValue
End Property

Property Get MultiSelectIndex_Data(Indexs As Long) As Long
    MultiSelectIndex_Data = Data_IndexMultiSelect.Item(Indexs)
End Property

Property Get MultiSelectIndex_Count() As Long
    MultiSelectIndex_Count = Count_IndexMultiSelect
End Property


Private Sub RemoveXY(Optional ByVal IndexX As Long = -1, Optional ByVal IndexY As Long = -1)
Dim TmpGridXYData() As TypeGridXYData
Dim OpX As Boolean, OpY As Boolean
Dim VpX As Integer, VpY As Integer
Dim CountFixX As Long, CountSellX As Long
Dim CountFixY As Long, CountSellY As Long

        TmpGridXYData() = GridXYData()
        
        If IndexX > -1 Then
            IndexX = IndexX + 1
            OpX = True
        Else
            IndexX = 0
        End If
        
        If IndexY > -1 Or MultiSelectXY = 2 Then
            If IndexY < 0 And MultiSelectXY = 2 Then IndexY = On_IndexMultiSelectXY Else IndexY = IndexY + 1
            OpY = True
        Else
            IndexY = 0
        End If
        
        If OpX = True And MultiSelectXY = 1 Then
            ReDim GridXYData(SetNewGrid.GridXCount - Count_IndexMultiSelect, SetNewGrid.GridYCount - (1 + Abs(Int(OpY))))
        ElseIf OpY = True And MultiSelectXY = 2 Then
            ReDim GridXYData(SetNewGrid.GridXCount - (1 + Abs(Int(OpX))), SetNewGrid.GridYCount - Count_IndexMultiSelect)
        Else
            ReDim GridXYData(SetNewGrid.GridXCount - (1 + Abs(Int(OpX))), SetNewGrid.GridYCount - (1 + Abs(Int(OpY))))
        End If
        GridXYData() = TmpGridXYData()
        
        For X = IndexX To SetNewGrid.GridXCount - 1
            If OpX = True Then
                If MultiSelectXY = 1 Then
                    If IndexMultiSelectXY(Y) = True Then CountFixX = CountFixX + 1 Else GridX(X - CountFixX) = GridX(X)
                Else
                    GridX(X - 1) = GridX(X)
                End If
            End If
            
            If MultiSelectXY = 2 Then CountSellY = 0
            For Y = IndexY To SetNewGrid.GridYCount - 1
                If OpY = True And X = IndexX Then
                    If MultiSelectXY = 2 Then
                        If IndexMultiSelectXY(Y) = True Then CountFixY = CountFixY + 1 Else GridY(Y - CountFixY) = GridY(Y)
                    Else
                        GridY(Y - 1) = GridY(Y)
                    End If
                End If
                
                If OpX = True Then
                    If MultiSelectXY = 1 Then
                        If IndexMultiSelectXY(Y) = True Then CountSellX = CountSellX + 1 Else GridXYData(X - CountSellX, Y) = TmpGridXYData(X, Y)
                    Else
                        GridXYData(X - 1, Y) = TmpGridXYData(X, Y)
                    End If
                End If
                
                If OpY = True Then
                    If MultiSelectXY = 2 Then
                        If IndexMultiSelectXY(Y) = True Then CountSellY = CountSellY + 1 Else GridXYData(X, Y - CountSellY) = TmpGridXYData(X, Y)
                    Else
                        GridXYData(X, Y - 1) = TmpGridXYData(X, Y)
                    End If
                End If
            Next Y
        Next X
        
        If OpX = True Then
            If MultiSelectXY = 1 Then
                SetNewGrid.GridXCount = SetNewGrid.GridXCount - Count_IndexMultiSelect
            Else
                SetNewGrid.GridXCount = SetNewGrid.GridXCount - 1
            End If
            
            ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
            Hsc.Max = SetNewGrid.GridXCount - 1
        End If
        
        If OpY = True Then
            If MultiSelectXY = 2 Then
                SetNewGrid.GridYCount = SetNewGrid.GridYCount - Count_IndexMultiSelect
            Else
                SetNewGrid.GridYCount = SetNewGrid.GridYCount - 1
            End If
            
            ReDim Preserve GridY(SetNewGrid.GridYCount - 1)
            Vsc.Max = SetNewGrid.GridYCount - 1
        End If
    
'    Command3.Caption = IndexX
End Sub

Sub RemoveColumns(Optional Index As Long = -1)
    RemoveXY Index
End Sub

Sub RemoveRows(Optional Index As Long = -1)
    RemoveXY , Index
End Sub


Sub ScrollHRL(Optional ByVal IndexsX As Long = -1, Optional ByVal IndexsY As Long = -1, Optional AutosX As Boolean, Optional AutosY As Boolean)
Dim iX As Long, iY As Long, OverGrid As Long
    
    If IndexsX > -1 Then
        If IndexsX >= GridRightX And AutosX = True Then
            For iX = IndexsX To 0 Step -1
                If GridX(ShowFixCol(iX)).GWidthDefault = False Then GridX(ShowFixCol(iX)).GridWidth = SetNewGrid.GridSize.SellWidth_Def
                OverGrid = OverGrid + GridX(ShowFixCol(iX)).GridWidth
                
                If OverGrid > Picture1.ScaleWidth - SetNewGrid.GridSize.GDRangeX Then
                    IndexsX = iX + 1
                    Exit For
                End If
            Next iX
        End If
    End If
    OverGrid = 0
    If IndexsY > -1 Then
        If IndexsY >= GridDownY And AutosY = True Then
            For iY = IndexsY To 0 Step -1
                If GridY(iY).GHeightDefault = False Then GridY(iY).GridHeight = SetNewGrid.GridSize.SellHeight_Def
                OverGrid = OverGrid + GridY(iY).GridHeight
                
                If OverGrid > Picture1.ScaleHeight - SetNewGrid.GridSize.GDRangeY Then
                    IndexsY = iY + 1
                    Exit For
                End If
            Next iY
        End If
    End If
    
    If IndexsX > -1 Or IndexsY > -1 Then
        TypeDragMove = True
        'For iX = 0 To UBound(ShowFixCol()) - 1
        '    If ShowFixCol(iX) = IndexsX Then
        '        IndexsX = iX
        '    Exit For
        '    End If
        'Next iX
        If IndexsX > -1 And Hsc.Max > 0 Then Hsc.Value = IndexsX
        If IndexsY > -1 And Vsc.Max > 0 Then Vsc.Value = IndexsY
        TypeDragMove = False

        ClickScroll = False
        
        GridXY Picture1, Hsc.Value, Vsc.Value
        Picture1.Picture = Picture1.Image
            If CheckType(GridType, 0) = False Then _
            DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
'On On            DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    End If
End Sub

Sub Tmp_ScrollHRL(Optional ByVal IndexsX As Long = -1, Optional ByVal IndexsY As Long = -1, Optional AutosX As Boolean, Optional AutosY As Boolean)
Dim iX As Long, iY As Long, OverGrid As Long
    
ShowFixCol
    If IndexsX > -1 Then
        If IndexsX >= GridRightX And AutosX = True Then
            For iX = IndexsX To 0 Step -1
                If GridX(iX).GWidthDefault = False Then GridX(iX).GridWidth = SetNewGrid.GridSize.SellWidth_Def
                OverGrid = OverGrid + GridX(iX).GridWidth
                
                If OverGrid > Picture1.ScaleWidth - SetNewGrid.GridSize.GDRangeX Then
                    IndexsX = iX + 1
                    Exit For
                End If
            Next iX
        End If
    End If
    OverGrid = 0
    If IndexsY > -1 Then
        If IndexsY >= GridDownY And AutosY = True Then
            For iY = IndexsY To 0 Step -1
                If GridY(iY).GHeightDefault = False Then GridY(iY).GridHeight = SetNewGrid.GridSize.SellHeight_Def
                OverGrid = OverGrid + GridY(iY).GridHeight
                
                If OverGrid > Picture1.ScaleHeight - SetNewGrid.GridSize.GDRangeY Then
                    IndexsY = iY + 1
                    Exit For
                End If
            Next iY
        End If
    End If
    
    If IndexsX > -1 Or IndexsY > -1 Then
        TypeDragMove = True
        If IndexsX > -1 Then Hsc.Value = IndexsX
        If IndexsY > -1 Then Vsc.Value = IndexsY
        TypeDragMove = False

        ClickScroll = False
        
        GridXY Picture1, Hsc.Value, Vsc.Value
        Picture1.Picture = Picture1.Image
            If CheckType(GridType, 0) = False Then _
            DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    End If
End Sub


'Sub HScroll()
    
'End Sub


'Property Get SellText(ByVal IndexX As Long, ByVal IndexY As Long) As String
'    SellText = GridXYData(IndexX, IndexY).GridXYValue
'End Property
'Property Let SellText(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As String)
'    GridXYData(IndexX, IndexY).GridXYValue = NewValue
'End Property


'Property Get Tag(ByVal IndexX As Long, ByVal IndexY As Long) As String
'    Tag = GridXYData(IndexX, IndexY).GridTag
'End Property
'Property Let Tag(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As String)
'    GridXYData(IndexX, IndexY).GridTag = NewValue
'End Property

Sub SetingPointer(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) 'Belom ada coment
    Pxy.Px1 = X1
    Pxy.Px2 = X2
    Pxy.Py1 = Y1
    Pxy.Py2 = Y2
End Sub

'Poerty Text -------------------------------------------------------------------------
Sub SetText()
    Set ObjcText = Text1
End Sub

'Sub ShowText(IndexX As Long, IndexY As Long, Optional ClearText As Boolean)
Sub ShowText(IndexX As Long, IndexY As Long, Optional IndexXScroll As Long = -1, Optional IndexYScroll As Long = -1, Optional ClearText As Boolean, Optional NoSrolls As Boolean)

    'Text1.Visible = False
    If NoSrolls = False Then
    If IndexXScroll = -1 Then IndexXScroll = IndexX
    If IndexYScroll = -1 Then IndexYScroll = IndexY
    
    If GridX(ShowFixCol(IndexXScroll)).GWidthDefault = False Then GridX(ShowFixCol(IndexXScroll)).GridWidth = SetNewGrid.GridSize.SellWidth_Def
    If GridY(IndexYScroll).GHeightDefault = False Then GridY(IndexYScroll).GridHeight = SetNewGrid.GridSize.SellHeight_Def

        If IndexXScroll >= GridRightX And _
        IndexYScroll >= GridDownY Then
            ScrollHRL IndexXScroll, IndexYScroll, True, True
        ElseIf IndexXScroll < GridLeftX And IndexYScroll < GridUpY Then
            ScrollHRL IndexXScroll, IndexYScroll
        ElseIf IndexXScroll < GridLeftX Then
            ScrollHRL IndexXScroll
        ElseIf IndexXScroll >= GridRightX Then
            ScrollHRL IndexXScroll, , True
        ElseIf IndexYScroll < GridUpY Then
            ScrollHRL , IndexYScroll
        ElseIf IndexYScroll >= GridDownY Then
    '        If GridY(IndexYScroll).GridTop + GridY(IndexYScroll).GridHeight > Picture1.ScaleHeight Then ScrollHRL , IndexYScroll, , True
            ScrollHRL , IndexYScroll, , True
        End If
    End If
    
    AllText Text1, SellText(ShowFixCol(IndexX), IndexY), SellLeft(ShowFixCol(IndexX)), SellTop(IndexY), SellWidth(ShowFixCol(IndexX)), SellHeight_Def
 
    XPointerIndex = IndexX '?
    YPointerIndex = IndexY '?
    XIndexPointerText1 = IndexX
    YIndexPointerText1 = IndexY
    
    If ClearText = True Then Text1.Text = ""
'    ToolTipText IndexX, IndexY

'    Text1.ToolTipText = "OPOPOggggggggggggggggggggggggggg" & vbCrLf & "kkkkkkkkkkk"
'    tmpPicture1.Picture = Picture1.Picture
'    Picture1.Cls
    'Picture1.Picture = Picture1.Image
'    Picture1.Line (0, 0)-(Int(Rnd * 100), 100)
End Sub

Sub TMP_ShowText(IndexX As Long, IndexY As Long, Optional IndexXScroll As Long = -1, Optional IndexYScroll As Long = -1, Optional ClearText As Boolean, Optional NoSrolls As Boolean)

    'Text1.Visible = False
    If NoSrolls = False Then
    If IndexXScroll = -1 Then IndexXScroll = IndexX
    If IndexYScroll = -1 Then IndexYScroll = IndexY
    
    If GridX(IndexXScroll).GWidthDefault = False Then GridX(IndexXScroll).GridWidth = SetNewGrid.GridSize.SellWidth_Def
    If GridY(IndexYScroll).GHeightDefault = False Then GridY(IndexYScroll).GridHeight = SetNewGrid.GridSize.SellHeight_Def

        If IndexXScroll >= GridRightX And _
        IndexYScroll >= GridDownY Then
            ScrollHRL IndexXScroll, IndexYScroll, True, True
        ElseIf IndexXScroll < GridLeftX And IndexYScroll < GridUpY Then
            ScrollHRL IndexXScroll, IndexYScroll
        ElseIf IndexXScroll < GridLeftX Then
            ScrollHRL IndexXScroll
        ElseIf IndexXScroll >= GridRightX Then
            ScrollHRL IndexXScroll, , True
        ElseIf IndexYScroll < GridUpY Then
            ScrollHRL , IndexYScroll
        ElseIf IndexYScroll >= GridDownY Then
    '        If GridY(IndexYScroll).GridTop + GridY(IndexYScroll).GridHeight > Picture1.ScaleHeight Then ScrollHRL , IndexYScroll, , True
            ScrollHRL , IndexYScroll, , True
        End If
    End If
    
    AllText Text1, SellText(IndexX, IndexY), SellLeft(IndexX), SellTop(IndexY), SellWidth(IndexX), SellHeight_Def
 
    XPointerIndex = IndexX '?
    YPointerIndex = IndexY '?
    XIndexPointerText1 = IndexX
    YIndexPointerText1 = IndexY
    
    If ClearText = True Then Text1.Text = ""
'    ToolTipText IndexX, IndexY

'    Text1.ToolTipText = "OPOPOggggggggggggggggggggggggggg" & vbCrLf & "kkkkkkkkkkk"
'    tmpPicture1.Picture = Picture1.Picture
'    Picture1.Cls
    'Picture1.Picture = Picture1.Image
'    Picture1.Line (0, 0)-(Int(Rnd * 100), 100)
End Sub

Function SetInputText(Optional Texts As String) As String
    If Texts <> "" Then Text1.Text = Texts
    SetInputText = Text1.Text
    
    'Text1.SelStart = Len(Texts)
    'Text1.SelLength = 2
End Function

Property Get SetTextVisible() As Boolean
    SetTextVisible = Text1.Visible
End Property
Property Let SetTextVisible(ByVal NewValue As Boolean)
    Text1.Visible = NewValue
    If Text1.Visible = False Then Picture1.SetFocus
End Property


'Function SetTextVisible(Optional nVisible As Boolean)
'    Text1.Visible = nVisible
'    Picture1.SetFocus
'End Function
'Function GetTextVisible() As Boolean
'    GetTextVisible = Text1.Visible
'End Function



'End Poperty Text-------------------------------------------------------------------------

'Poerty List -------------------------------------------------------------------------
Sub AddList(IndexGridX As Long, IndexGridY As Long, nCount As Integer)
Dim iY As Integer, nCountFrist As Integer

    ReDim Preserve GridXYData(IndexGridX, IndexGridY).GridSubType.ContGList(nCount)

    If GridXYData(IndexGridX, IndexGridY).GridSubType.CountContGList = 0 Then nCountFrist = 0 Else nCountFrist = nCount - GridXYData(IndexGridX, IndexGridY).GridSubType.CountContGList
    For iY = nCountFrist To nCount
        GridXYData(IndexGridX, IndexGridY).GridSubType.ContGList(iY).GLHeadColor = &HFF8080
        GridXYData(IndexGridX, IndexGridY).GridSubType.ContGList(iY).GLHeadColorClik = &HFFC0C0
    Next iY
    
    If nCount > 0 Then GridXYData(IndexGridX, IndexGridY).GridSubType.CountContGList = nCount
    GridXYData(IndexGridX, IndexGridY).GridSubType.CountContGList = GridXYData(IndexGridX, IndexGridY).GridSubType.CountContGList + 1

'    If nCount > 0 Then CountGridList = nCount
'    ReDim Preserve GridList(CountGridList)
'    CountGridList = CountGridList + 1
End Sub
Sub ClearList(IndexGridX As Long, IndexGridY As Long)
    Erase GridXYData(IndexGridX, IndexGridY).GridSubType.ContGList()
    GridXYData(IndexGridX, IndexGridY).GridSubType.CountContGList = 0
End Sub

Sub AddListText(IndexGridX As Long, IndexGridY As Long, IndexList As Long, nCount As Long)
    GridXYData(IndexGridX, IndexGridY).GridSubType.ContGList(IndexList).Add nCount
End Sub
Sub ClearListText(IndexX As Long, IndexY As Long, IndexList As Integer)
    GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).Clear_GLText
End Sub


Property Get ListHeadColorClik(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer) As Long
    ListHeadColor = GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLHeadColorClik
End Property
Property Let ListHeadColorClik(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer, ByVal NewValue As Long)
    GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLHeadColorClik = NewValue
End Property

Property Get ListHeadColor(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer) As Long
    ListHeadColor = GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLHeadColor
End Property
Property Let ListHeadColor(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer, ByVal NewValue As Long)
    GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLHeadColor = NewValue
End Property

Property Get ListCaption(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer) As String
    ListCaption = GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLCaption
End Property
Property Let ListCaption(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer, ByVal NewValue As String)
    GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLCaption = NewValue
End Property

Property Get ListShowCount(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    ListShowCount = GridXYData(IndexX, IndexY).GridSubType.GLShowCount
End Property
Property Let ListShowCount(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
    GridXYData(IndexX, IndexY).GridSubType.GLShowCount = NewValue
End Property

Property Get ListHeight(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    ListHeight = GridXYData(IndexX, IndexY).GridSubType.GLHeight
End Property
Property Let ListHeight(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
    GridXYData(IndexX, IndexY).GridSubType.GLHeight = NewValue
End Property

Property Get ListText(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer, ByVal IndexListText As Long) As String
    ListText = GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLText(IndexListText)
End Property
Property Let ListText(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer, ByVal IndexListText As Long, ByVal NewValue As String)
    GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLText(IndexListText) = NewValue
End Property

Property Get ListIndex(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    ListIndex = GridXYData(IndexX, IndexY).GridSubType.PointerListHead
End Property
Property Let ListIndex(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
    GridXYData(IndexX, IndexY).GridSubType.PointerListHead = NewValue
End Property

Property Get ListCount(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    ListCount = GridXYData(IndexX, IndexY).GridSubType.CountContGList
End Property

Property Get ListIndexText(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer) As Integer
    ListIndexText = GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLPointer
End Property
Property Let ListIndexText(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Integer, ByVal NewValue As Integer)
    IndexListS IndexX, IndexY, IndexList, NewValue
End Property

Property Get ListCountText(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Long) As Integer
    ListCountText = GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLCount
End Property

Property Get ListType(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    ListType = GridXYData(IndexX, IndexY).GridSubType.TypeControl
End Property
Property Let ListType(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
    GridXYData(IndexX, IndexY).GridSubType.TypeControl = NewValue
End Property

Property Get ListRange(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Long) As Integer
    ListRange = GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLRange
End Property
Property Let ListRange(ByVal IndexX As Long, ByVal IndexY As Long, ByVal IndexList As Long, ByVal NewValue As Integer)
    GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList).GLRange = NewValue
End Property

'End Poperty List-------------------------------------------------------------------------

Property Get TypeControl(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    TypeControl = GridXYData(IndexX, IndexY).GridSubType.TypeControl
End Property
Property Let TypeControl(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
    GridXYData(IndexX, IndexY).GridSubType.TypeControl = NewValue
End Property


'Property Get AddList FixdRowsColStyl(ByVal Index As Long) As Integer
'    AddList = GridY(Index).GridStyle
'End Property
'Property Let FixdRowsColStyl(ByVal Index As Long, ByVal NewValue As Integer)
'    GridY(Index).GridStyle = NewValue
'End Property

Property Get SellTag(ByVal IndexX As Long, ByVal IndexY As Long) As String
    SellTag = GridXYData(IndexX, IndexY).GridTag
End Property
Property Let SellTag(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As String)
    GridXYData(IndexX, IndexY).GridTag = NewValue
End Property

Property Get SellTagSub(ByVal IndexX As Long, ByVal IndexY As Long) As String
    SellTagSub = GridXYData(IndexX, IndexY).GridTagSub
End Property
Property Let SellTagSub(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As String)
    GridXYData(IndexX, IndexY).GridTagSub = NewValue
End Property



Property Get PictureSubGridY() As StdPicture
    Set PictureSubGridY = tmpPictureSubGridY
End Property
Property Let PictureSubGridY(ByVal NewValue As StdPicture)
    Set tmpPictureSubGridY = NewValue
End Property

Property Get PictureGrid() As StdPicture
'    Set tmpPictureGrid = tmpPictureGrid
     Set PictureGrid = tmpPictureGrid
End Property
Property Set PictureGrid(ByVal NewValue As StdPicture)
    Set tmpPictureGrid = NewValue
End Property
Property Let PictureGrid(ByVal NewValue As StdPicture)
    Set tmpPictureGrid = NewValue
End Property

'Grid ***********************************************************************************************************************************
Property Get GridLeft() As Integer
    GridLeft = GridLeftX
End Property
Property Let GridLeft(ByVal NewValue As Integer)
    Hsc.Value = NewValue
End Property

Property Get GridRight() As Integer
    GridRight = GridRightX
End Property

Property Get GridUp() As Integer
    GridUp = GridUpY
End Property
Property Let GridUp(ByVal NewValue As Integer)
    GridUpY = NewValue
    Vsc.Value = GridUpY 'No clear

End Property

Property Get GridDown() As Integer
    GridDown = GridDownY
End Property
'***********************************************************************************************************************************

'Sell Fixed ***********************************************************************************************************************************
'Sell Fixed Move ---------------------------------------------------------------------------------------------------------------------------------
Property Get FixedWidth() As Integer
    FixedWidth = SetNewGrid.GridSize.GDRangeX
End Property
Property Let FixedWidth(ByVal NewValue As Integer)
    SetNewGrid.GridSize.GDRangeX = NewValue
End Property

Property Get FixedHeight() As Integer
    FixedHeight = SetNewGrid.GridSize.GDRangeY
End Property
Property Let FixedHeight(ByVal NewValue As Integer)
    SetNewGrid.GridSize.GDRangeY = NewValue
End Property
'---------------------------------------------------------------------------------------------------------------------------------
'Sell Fixed SellText---------------------------------------------------------------------------------------------------------------------------------
Property Get FixedColsText(ByVal Index As Long) As String
    FixedColsText = GridX(Index).GridValue
End Property
Property Let FixedColsText(ByVal Index As Long, ByVal NewValue As String)
    GridX(Index).GridValue = NewValue
End Property


Property Get FixedRowsText(ByVal Index As Long) As String
    FixedRowsText = GridY(Index).GridValue
End Property
Property Let FixedRowsText(ByVal Index As Long, ByVal NewValue As String)
    GridY(Index).GridValue = NewValue
End Property
'---------------------------------------------------------------------------------------------------------------------------------
'Sell Fixed Default Color---------------------------------------------------------------------------------------------------------------------------------
Property Get FixdColnDefColStyl() As Integer
    FixdColnDefColStyl = SetNewGrid.GridStyleX
End Property
Property Let FixdColnDefColStyl(ByVal NewValue As Integer)
    SetNewGrid.GridStyleX = NewValue
End Property

Property Get FixdColnDefColBck0() As OLE_COLOR
    FixdColnDefColBck0 = SetNewGrid.GridLenkapX.GridBackColGra(0)
End Property
Property Let FixdColnDefColBck0(ByVal NewValue As OLE_COLOR)
    SetNewGrid.GridLenkapX.GridBackColGra(0) = NewValue
End Property
Property Get FixdColnDefColBck1() As OLE_COLOR
    FixdColnDefColBck1 = SetNewGrid.GridLenkapX.GridBackColGra(1)
End Property
Property Let FixdColnDefColBck1(ByVal NewValue As OLE_COLOR)
    SetNewGrid.GridLenkapX.GridBackColGra(1) = NewValue
End Property
Sub FixdColnDefColGradn(Color1 As OLE_COLOR, Color2 As OLE_COLOR)
    SetNewGrid.GridLenkapX.GridBackColGra(0) = Color1
    SetNewGrid.GridLenkapX.GridBackColGra(1) = Color2
End Sub
'
'
Property Get FixdRowsDefColStyl() As Integer
    FixdRowsDefColStyl = SetNewGrid.GridStyleY
End Property
Property Let FixdRowsDefColStyl(ByVal NewValue As Integer)
    SetNewGrid.GridStyleY = NewValue
End Property

Property Get FixdRowsDefColBck0() As OLE_COLOR
    FixdRowsDefColBck0 = SetNewGrid.GridLenkapY.GridBackColGra(0)
End Property
Property Let FixdRowsDefColBck0(ByVal NewValue As OLE_COLOR)
    SetNewGrid.GridLenkapY.GridBackColGra(0) = NewValue
End Property
Property Get FixdRowsDefColBck1() As OLE_COLOR
    FixdRowsDefColBck1 = SetNewGrid.GridLenkapY.GridBackColGra(1)
End Property
Property Let FixdRowsDefColBck1(ByVal NewValue As OLE_COLOR)
    SetNewGrid.GridLenkapY.GridBackColGra(1) = NewValue
End Property
Sub FixdRowsDefColGradn(Color1 As OLE_COLOR, Color2 As OLE_COLOR)
    SetNewGrid.GridLenkapY.GridBackColGra(0) = Color1
    SetNewGrid.GridLenkapY.GridBackColGra(1) = Color2
End Sub
'---------------------------------------------------------------------------------------------------------------------------------
'Sell Fixed Color---------------------------------------------------------------------------------------------------------------------------------
Property Get FixdColnColStyl(ByVal Index As Long) As Integer
    FixdColnColStyl = GridX(Index).GridStyle
End Property
Property Let FixdColnColStyl(ByVal Index As Long, ByVal NewValue As Integer)
    GridX(Index).GridStyle = NewValue
End Property

Property Get FixdColnTag(ByVal Index As Long) As String
    FixdColnTag = GridX(Index).Tag
End Property
Property Let FixdColnTag(ByVal Index As Long, ByVal NewValue As String)
    GridX(Index).Tag = NewValue
End Property

'New
'    ConHeadSellOnFixCol = 3
'    ReDim Preserve HeadSellOnFixCol(ConHeadSellOnFixCol - 1)
'    HeadSellOnFixCol(0) = 0 '.IndexFixCol = 0
'Sub Add(Optional ByVal VX As Long, Optional ByVal VY As Long, Optional NewType As TypeAdd = -1)
Sub Add_FixdColnShowHead(ByVal Index As Integer)
Dim iX As Integer, TmpIndex As Integer, TmpHead As Integer

'MsgBox GridX(5).GridIndexHead ShowOnConHeadSell
If GridX(Index).GridIndexHead = 0 Then
'    Hsc.Max = Hsc.Max - 1
    ConHeadSellOnFixCol = ConHeadSellOnFixCol + 1
    ReDim Preserve HeadSellOnFixCol(ConHeadSellOnFixCol - 1)

    If ConHeadSellOnFixCol > 1 And 8 = 8 Then
        If Index < HeadSellOnFixCol(ConHeadSellOnFixCol - 2) Then
            TmpIndex = -1
            For iX = 0 To ConHeadSellOnFixCol - 1
                If Index < HeadSellOnFixCol(iX) Then
                    TmpIndex = iX
                    Exit For
                End If
            Next iX
            For iX = ConHeadSellOnFixCol - 2 To TmpIndex Step -1
                GridX(HeadSellOnFixCol(iX)).GridIndexHead = iX + 2
                HeadSellOnFixCol(iX + 1) = HeadSellOnFixCol(iX)
            Next iX
            HeadSellOnFixCol(TmpIndex) = Index
            GridX(HeadSellOnFixCol(TmpIndex)).GridIndexHead = TmpIndex + 1
        Else
            HeadSellOnFixCol(ConHeadSellOnFixCol - 1) = Index
            GridX(Index).GridIndexHead = ConHeadSellOnFixCol
        End If
    Else
        HeadSellOnFixCol(ConHeadSellOnFixCol - 1) = Index
        GridX(Index).GridIndexHead = ConHeadSellOnFixCol
    End If
End If
HscHead.Max = ConHeadSellOnFixCol - 1

HeadOnHSC

'If Index = 1 Then
'For iX = 0 To ConHeadSellOnFixCol - 1
'    MsgBox iX & "." & HeadSellOnFixCol(iX) & " "
'MsgBox GridX(HeadSellOnFixCol(iX)).GridIndexHead
'Next iX
'End If 'GridIndexHead

End Sub 'ConHeadSellOnFixCol

Property Get FixdColnShowHead_Count() As Integer
    FixdColnShowHead_Count = ConHeadSellOnFixCol
End Property

Property Get FixdColnShowHead_RunCount() As Integer
    FixdColnShowHead_RunCount = ShowOnConHeadSell
End Property

Property Get FixdColnShowHead(ByVal Index As Long) As Integer
    FixdColnShowHead = HeadSellOnFixCol(Index)
End Property
Property Let FixdColnShowHead(ByVal Index As Long, ByVal NewValue As Integer)
    HeadSellOnFixCol(Index) = NewValue
End Property

Property Get FixdColnIndex_ShowHead(ByVal Index As Long) As Integer
    FixdColnIndex_ShowHead = GridX(Index).GridIndexHead
End Property

Property Get FixdColnRangeEndHead_Pos() As Integer
    FixdColnRangeEndHead_Pos = RangeEndHead_Pos
End Property
Property Let FixdColnRangeEndHead_Pos(ByVal NewValue As Integer)
    RangeEndHead_Pos = NewValue
    RangeEndHead = RangeEndHead_Pos
End Property

Property Get FixdColnRangeEndHead_PosAuto() As Integer
    FixdColnRangeEndHead_PosAuto = RangeEndHead_PosAuto
End Property
Property Let FixdColnRangeEndHead_PosAuto(ByVal NewValue As Integer)
    RangeEndHead_PosAuto = NewValue
End Property

Property Get FixdColnHeadCountAuto() As Boolean
    FixdColnHeadCountAuto = HeadCountAuto
End Property
Property Let FixdColnHeadCountAuto(ByVal NewValue As Boolean)
    HeadCountAuto = NewValue
    
    If HeadCountAuto = False Then RangeEndHead = RangeEndHead_Pos
End Property

'HeadCountAuto = False

Property Get FixdColnIndex_OnReal(ByVal Index As Long) As Integer
    If GridX(Index).GridRealOnPosisi = -1 Then GridX(Index).GridRealOnPosisi = Index
    FixdColnIndex_OnReal = GridX(Index).GridRealOnPosisi
End Property

Property Get FixdColnIndex_Real(ByVal Index As Long) As Integer
    If ShowFixCol(0) <> Index And GridX(Index).GridRealPosisi <= 0 Then GridX(Index).GridRealPosisi = Index
    FixdColnIndex_Real = GridX(Index).GridRealPosisi
End Property

Property Get FixdHideCountCol() As Integer
    FixdHideCountCol = HideCountFixCol
End Property

Property Get FixdColRangeCountHead() As Integer
    FixdColRangeCountHead = RangeEndHead
End Property
Property Let FixdColRangeCountHead(ByVal NewValue As Integer)
    RangeEndHead = NewValue
End Property

Property Get FixColIndexHead(ByVal Index As Long) As Integer
    If ShowOnConHeadSell > 0 And Index < 0 Then
        FixColIndexHead = HeadSellOnFixCol(Abs(Index) - 1)
    Else
        FixColIndexHead = Index
    End If
End Property
Property Let FixColIndexHead(ByVal Index As Long, ByVal NewValue As Integer)  'As Integer
    If ShowOnConHeadSell > 0 And Index < 0 Then HeadSellOnFixCol(Abs(Index) - 1) = NewValue
End Property

Property Get FixdColnIndexShow_Count() As Integer
    FixdColnIndexShow_Count = ConShowFixCol 'ShowOnConHeadSell
'    FixdColnIndexShow_Count = ShowOnConHeadSell
End Property

Property Get FixdColnIndex(ByVal Index As Long) As Integer
    FixdColnIndex = ShowFixCol(FixColIndexHead(Index))
End Property
Property Let FixdColnIndex(ByVal Index As Long, ByVal NewValue As Integer)
Dim iX As Integer, TMPShowFixCol As Integer, TMPRealPosisi As Integer
Dim TMPGridRealOnPosisi As Integer, TmpGridIndexHead As Integer
Dim HADGridRealOnPosisi As Integer
'GridIndexHead ConShowFixCol

'MsgBox GridX(ShowFixCol(Index)).GridIndexHead & " " & GridX(ShowFixCol(NewValue)).GridIndexHead


    GridX(ShowFixCol(Index)).GridLeft = -1
    
    TMPShowFixCol = ShowFixCol(Index)
    TMPRealPosisi = GridX(ShowFixCol(NewValue)).GridRealPosisi
    
    ShowFixCol(Index) = ShowFixCol(NewValue)
    If TMPRealPosisi = 0 Then TMPRealPosisi = NewValue 'ShowFixCol(NewValue)
    ShowFixCol(TMPRealPosisi) = TMPShowFixCol 'TMPShowFixCol

    'Head
    TmpGridIndexHead = GridX(ShowFixCol(Index)).GridIndexHead
    HADGridRealOnPosisi = GridX(ShowFixCol(Index)).GridRealPosisi
    If TmpGridIndexHead = 0 Then
        TmpGridIndexHead = GridX(ShowFixCol(NewValue)).GridIndexHead
        HADGridRealOnPosisi = GridX(ShowFixCol(NewValue)).GridRealPosisi
    End If 'Head Continue To 1
    
    GridX(ShowFixCol(NewValue)).GridRealPosisi = NewValue
    GridX(ShowFixCol(Index)).GridRealPosisi = Index
    
    If GridX(ShowFixCol(NewValue)).GridRealOnPosisi = -1 Then GridX(ShowFixCol(NewValue)).GridRealOnPosisi = ShowFixCol(NewValue)
    If GridX(ShowFixCol(Index)).GridRealOnPosisi = -1 Then GridX(ShowFixCol(Index)).GridRealOnPosisi = ShowFixCol(Index)
    TMPGridRealOnPosisi = GridX(ShowFixCol(NewValue)).GridRealOnPosisi
    GridX(ShowFixCol(NewValue)).GridRealOnPosisi = GridX(ShowFixCol(Index)).GridRealOnPosisi
    GridX(ShowFixCol(Index)).GridRealOnPosisi = TMPGridRealOnPosisi
    
    If GridX(ShowFixCol(Index)).GridIndexHead <> 0 And GridX(ShowFixCol(NewValue)).GridIndexHead <> 0 Then
        TMPGridRealOnPosisi = GridX(ShowFixCol(NewValue)).GridIndexHead
        GridX(ShowFixCol(NewValue)).GridIndexHead = GridX(ShowFixCol(Index)).GridIndexHead
        GridX(ShowFixCol(Index)).GridIndexHead = TMPGridRealOnPosisi
            
        HeadSellOnFixCol(GridX(ShowFixCol(Index)).GridIndexHead - 1) = ShowFixCol(Index)
        HeadSellOnFixCol(GridX(ShowFixCol(NewValue)).GridIndexHead - 1) = ShowFixCol(NewValue)
    Else
        'Head 1
        If TmpGridIndexHead <> 0 Then
            If HADGridRealOnPosisi > GridX(HeadSellOnFixCol(TmpGridIndexHead - 1)).GridRealPosisi Then
                For iX = GridX(HeadSellOnFixCol(TmpGridIndexHead - 1)).GridIndexHead - 1 To 1 Step -1
                    If GridX(HeadSellOnFixCol(iX - 1)).GridRealOnPosisi = -1 Then
                        GridX(HeadSellOnFixCol(iX - 1)).GridRealOnPosisi = HeadSellOnFixCol(iX - 1)
                    End If
                    If GridX(HeadSellOnFixCol(TmpGridIndexHead - 1)).GridRealOnPosisi < GridX(HeadSellOnFixCol(iX - 1)).GridRealOnPosisi Then
                        TMPGridRealOnPosisi = HeadSellOnFixCol(iX)
                        HeadSellOnFixCol(iX) = HeadSellOnFixCol(iX - 1)
                        HeadSellOnFixCol(iX - 1) = TMPGridRealOnPosisi
                    
                        GridX(HeadSellOnFixCol(iX)).GridIndexHead = GridX(HeadSellOnFixCol(iX - 1)).GridIndexHead
                        GridX(HeadSellOnFixCol(iX - 1)).GridIndexHead = iX '- 1
                    
                        TmpGridIndexHead = TmpGridIndexHead - 1
                    End If
                Next iX '>>>>>>>>>>>>>>>>
            Else ' Lebih Besar
                For iX = GridX(HeadSellOnFixCol(TmpGridIndexHead - 1)).GridIndexHead To ConHeadSellOnFixCol - 1
                    If GridX(HeadSellOnFixCol(iX)).GridRealOnPosisi = -1 Then
                        GridX(HeadSellOnFixCol(iX)).GridRealOnPosisi = HeadSellOnFixCol(iX)
                    End If
                    If GridX(HeadSellOnFixCol(TmpGridIndexHead - 1)).GridRealOnPosisi > GridX(HeadSellOnFixCol(iX)).GridRealOnPosisi Then
                        TMPGridRealOnPosisi = HeadSellOnFixCol(iX - 1) 'GridX(HeadSellOnFixCol(iX - 1)).GridIndexHead
                        HeadSellOnFixCol(iX - 1) = HeadSellOnFixCol(iX)
                        HeadSellOnFixCol(iX) = TMPGridRealOnPosisi
                        
                        GridX(HeadSellOnFixCol(iX - 1)).GridIndexHead = GridX(HeadSellOnFixCol(iX)).GridIndexHead
                        GridX(HeadSellOnFixCol(iX)).GridIndexHead = iX + 1
                        
                        TmpGridIndexHead = TmpGridIndexHead + 1
                    Else
                        Exit For
                    End If
                Next iX
            End If
        End If
    End If
End Property

Property Get FixdColnVisible(ByVal Index As Long) As Boolean
    FixdColnVisible = GridX(Index).Visibles
End Property
Property Let FixdColnVisible(ByVal Index As Long, ByVal NewValue As Boolean)
Dim iX As Integer, TMPShowFixCol() As Integer
Dim IndexProses As Integer ', MyProses As Boolean
    
    If NewValue = True Then 'Hide
        
        If GridX(Index).Visibles = False Then
            IndexProses = FixdColnIndex_Real(Index)
            If GridX(Index).GridRealOnPosisi = -1 Then GridX(Index).GridRealOnPosisi = Index
            
            If IndexProses > -1 Then
                For iX = IndexProses To UBound(ShowFixCol()) - 1
                    ShowFixCol(iX) = ShowFixCol(iX + 1)
                    GridX(ShowFixCol(iX)).GridRealPosisi = iX '>
                Next iX
                ConShowFixCol = UBound(ShowFixCol()) - 1
                ReDim Preserve ShowFixCol(ConShowFixCol)
                Hsc.Max = Hsc.Max - 1
                HideCountFixCol = HideCountFixCol + 1
            End If
        End If
    Else 'Show
'        >>>>>>>>>>>>GridRealOnPosisi
        If GridX(Index).Visibles = True Then
            'IndexProses = GridX(Index).GridRealOnPosisi - (HideCountFixCol - 1)
            'If IndexProses <= 1 Then IndexProses = UBound(ShowFixCol()) + 1
            'MsgBox GridX(ShowFixCol(UBound(ShowFixCol()))).GridRealOnPosisi
            If GridX(ShowFixCol(UBound(ShowFixCol()))).GridRealOnPosisi = -1 Then GridX(ShowFixCol(UBound(ShowFixCol()))).GridRealOnPosisi = ShowFixCol(UBound(ShowFixCol()))
            If GridX(Index).GridRealOnPosisi > GridX(ShowFixCol(UBound(ShowFixCol()))).GridRealOnPosisi Then
                IndexProses = UBound(ShowFixCol()) + 1
            Else
                For iX = 0 To UBound(ShowFixCol()) - 0
                    If GridX(ShowFixCol(iX)).GridRealOnPosisi = -1 Then GridX(ShowFixCol(iX)).GridRealOnPosisi = ShowFixCol(iX)
                    If GridX(Index).GridRealOnPosisi < GridX(ShowFixCol(iX)).GridRealOnPosisi Then
                        IndexProses = iX
                    Exit For
                    Else
                        IndexProses = UBound(ShowFixCol()) + 1
                    End If
                Next iX
            End If
            
            TMPShowFixCol() = ShowFixCol()
            ConShowFixCol = UBound(ShowFixCol()) + 1
            ReDim Preserve ShowFixCol(ConShowFixCol)
            
            GridX(Index).GridRealPosisi = IndexProses
            ShowFixCol(IndexProses) = Index '>>>> this error by min
            For iX = IndexProses To UBound(ShowFixCol()) - 1
                ShowFixCol(iX + 1) = TMPShowFixCol(iX)
                GridX(ShowFixCol(iX + 1)).GridRealPosisi = iX + 1 '>
            Next iX
            Hsc.Max = Hsc.Max + 1
            HideCountFixCol = HideCountFixCol - 1
            
        End If
    End If
    GridX(Index).Visibles = NewValue
End Property
'End New

Property Get FixdColnColBck0(ByVal Index As Long) As OLE_COLOR
    FixdColnColBck0 = GridX(Index).GridColGra(0)
End Property
Property Let FixdColnColBck0(ByVal Index As Long, ByVal NewValue As OLE_COLOR)
    GridX(Index).GridColGra(0) = NewValue
End Property
Property Get FixdColnColBck1(ByVal Index As Long) As OLE_COLOR
    FixdColnColBck1 = GridX(Index).GridColGra(1)
End Property
Property Let FixdColnColBck1(ByVal Index As Long, ByVal NewValue As OLE_COLOR)
    GridX(Index).GridColGra(1) = NewValue
End Property
Sub FixdColnColGradn(Index As Long, Color1 As OLE_COLOR, Color2 As OLE_COLOR)
    GridX(Index).GridColGra(0) = Color1
    GridX(Index).GridColGra(1) = Color2
End Sub

Property Get FixdColnForeColor() As Long
    FixdColnForeColor = SetNewGrid.GridLenkapX.GridForeColor
End Property
Property Let FixdColnForeColor(ByVal NewValue As Long)
    SetNewGrid.GridLenkapX.GridForeColor = NewValue
End Property
'
'
Property Get FixdRowsColStyl(ByVal Index As Long) As Integer
    FixdRowsColStyl = GridY(Index).GridStyle
End Property
Property Let FixdRowsColStyl(ByVal Index As Long, ByVal NewValue As Integer)
    GridY(Index).GridStyle = NewValue
End Property

Property Get FixdRowsColBck0(ByVal Index As Long) As OLE_COLOR
    FixdRowsColBck0 = GridY(Index).GridColGra(0)
End Property
Property Let FixdRowsColBck0(ByVal Index As Long, ByVal NewValue As OLE_COLOR)
    GridY(Index).GridColGra(0) = NewValue
End Property
Property Get FixdRowsColBck1(ByVal Index As Long) As OLE_COLOR
    FixdRowsColBck1 = GridY(Index).GridColGra(1)
End Property
Property Let FixdRowsColBck1(ByVal Index As Long, ByVal NewValue As OLE_COLOR)
    GridY(Index).GridColGra(1) = NewValue
End Property
Sub FixdRowsColGradn(Index As Long, Color1 As OLE_COLOR, Color2 As OLE_COLOR)
    GridY(Index).GridColGra(0) = Color1
    GridY(Index).GridColGra(1) = Color2
End Sub

Property Get FixdRowsForeColor() As Long
    FixdRowsForeColor = SetNewGrid.GridLenkapY.GridForeColor
End Property
Property Let FixdRowsForeColor(ByVal NewValue As Long)
    SetNewGrid.GridLenkapY.GridForeColor = NewValue
End Property

Property Get FixdRowTag(ByVal Index As Long) As String
    FixdRowTag = GridY(Index).Tag
End Property
Property Let FixdRowTag(ByVal Index As Long, ByVal NewValue As String)
    GridY(Index).Tag = NewValue
End Property

'---------------------------------------------------------------------------------------------------------------------------------
'***********************************************************************************************************************************

'Sell & Sell Sub***********************************************************************************************************************************
'Count---------------------------------------------------------------------------------------------------------------------------------
Property Get SellCountColumn() As Long
    SellCountColumn = SetNewGrid.GridXCount
End Property
Property Get SellCountRow() As Long
    SellCountRow = SetNewGrid.GridYCount
End Property
'---------------------------------------------------------------------------------------------------------------------------------
'Lebar Panjang---------------------------------------------------------------------------------------------------------------------------------
Property Get SellWidth_Def() As Integer
    SellWidth_Def = SetNewGrid.GridSize.SellWidth_Def
End Property
Property Let SellWidth_Def(ByVal NewValue As Integer)
    SetNewGrid.GridSize.SellWidth_Def = NewValue
End Property
Property Get SellHeight_Def() As Integer
    SellHeight_Def = SetNewGrid.GridSize.SellHeight_Def
End Property
Property Let SellHeight_Def(ByVal NewValue As Integer)
    SetNewGrid.GridSize.SellHeight_Def = NewValue
End Property


Property Get SellOpenWidth(ByVal Index As Long) As Boolean
    SellOpenWidth = GridX(Index).GWidthDefault
End Property
Property Let SellOpenWidth(ByVal Index As Long, ByVal MAValue As Boolean)
    GridX(Index).GWidthDefault = MAValue
End Property
Property Get SellOpenHeight(ByVal Index As Long) As Boolean
    SellOpenHeight = GridY(Index).GHeightDefault
End Property
Property Let SellOpenHeight(ByVal Index As Long, ByVal MAValue As Boolean)
    GridY(Index).GHeightDefault = MAValue
End Property

Property Get SellLeft(ByVal Index As Long, Optional ByVal WDefault As Boolean) As Single
    SellLeft = GridX(Index).GridLeft
End Property
'Property Let SellLeft(ByVal Index As Long, Optional ByVal WDefault As Boolean, ByVal NewValue As Single)
'    GridX(Index).GridLeft = NewValue
'        GridX(Index).GWidthDefault = WDefault
'End Property
Property Get SellTop(ByVal Index As Long, Optional ByVal HDefault As Boolean) As Single
    SellTop = GridY(Index).GridTop
End Property
'Property Let SellTop(ByVal Index As Long, Optional ByVal HDefault As Boolean, ByVal NewValue As Single)
'    GridY(Index).GridTop = NewValue
'        GridY(Index).GHidthDefault = HDefault
'End Property


Property Get SellWidth(ByVal Index As Long, Optional ByVal WDefault As Boolean) As Single
    SellWidth = GridX(Index).GridWidth
        'GridX(Index).GWidthDefault = WDefault
End Property
Property Let SellWidth(ByVal Index As Long, Optional ByVal WDefault As Boolean, ByVal NewValue As Single)
    GridX(Index).GridWidth = NewValue
        GridX(Index).GWidthDefault = WDefault
End Property

Property Get SellHeight(ByVal Index, Optional ByVal HDefault As Boolean) As Single
    'MsgBox UBound(GridY())
    SellHeight = GridY(Index).GridHeight
        'GridY(Index).GHidthDefault = HDefault
End Property
Property Let SellHeight(ByVal Index, Optional ByVal HDefault As Boolean, ByVal NewValue As Single)
    GridY(Index).GridHeight = SetNewGrid.GridSize.SellHeight_Def + NewValue
    GridY(Index).GHSave = NewValue
        If GridY(Index).GridHeight = SetNewGrid.GridSize.SellHeight_Def Then GridY(Index).GHeightDefault = False Else GridY(Index).GHeightDefault = True
End Property

Property Get SellGHSave(ByVal IndexY As Long) As Single
    SellGHSave = GridY(IndexY).GHSave
End Property
Property Let SellGHSave(ByVal IndexY As Long, ByVal NewValue As Single)
    GridY(IndexY).GHSave = NewValue
End Property
'---------------------------------------------------------------------------------------------------------------------------------
'Sell Font---------------------------------------------------------------------------------------------------------------------------------
Property Get SellText(ByVal IndexX As Long, ByVal IndexY As Long) As String
    SellText = GridXYData(IndexX, IndexY).GridXYValue
End Property
Property Let SellText(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As String)
    GridXYData(IndexX, IndexY).GridXYValue = NewValue
    If GridX(IndexX).Visibles = True Then Exit Property
    If GridX(IndexX).GridLeft <= 0 Then Exit Property
    
'MsgBox GridX(IndexX).GridOnIndexHead
'SetNewGrid.GridSize.GDRangeX = 25
'SetNewGrid.GridSize.GDRangeY = 20

'MsgBox GridX(25).GridLeft
'** Kemungkinan Error Jika Terjadi Pada Sell Yang Di Tiban Head
'    If (GridX(IndexX).GridIndexHead > 0 And GridX(IndexX).GridIndexHead <= ShowOnConHeadSell) Or (IndexX >= Hsc.Value And IndexX <= GridRightX) And (IndexY >= Vsc.Value And IndexY <= GridDownY) Then
    If ShowFixCol(0) <> IndexX And GridX(IndexX).GridRealPosisi <= 0 Then GridX(IndexX).GridRealPosisi = IndexX
    If ((GridX(IndexX).GridIndexHead > 0 And GridX(IndexX).GridIndexHead > HscHead.Value And GridX(IndexX).GridIndexHead <= ShowOnConHeadSell + HscHead.Value) And (GridX(IndexX).GridLeft >= SetNewGrid.GridSize.GDRangeX And GridX(IndexX).GridLeft <= Picture1.ScaleWidth) And (GridY(IndexY).GridTop >= SetNewGrid.GridSize.GDRangeY And GridY(IndexY).GridTop <= Picture1.ScaleHeight)) _
       Or (GridX(IndexX).GridOnIndexHead = 0 And (GridX(IndexX).GridRealPosisi >= Hsc.Value And GridX(IndexX).GridRealPosisi <= GridRightX) And (IndexY >= Vsc.Value And IndexY <= GridDownY)) Then

'MsgBox ShowOnConHeadSell
'HeadWidFixCol
'    If (GridX(IndexX).GridLeft >= SetNewGrid.GridSize.GDRangeX And GridX(IndexX).GridLeft <= Picture1.ScaleWidth) And (GridY(IndexY).GridTop >= SetNewGrid.GridSize.GDRangeY And GridY(IndexY).GridTop <= Picture1.ScaleHeight) Then
        Dim TmpCutGridX As Integer, TmpCutGridY As Integer, BColor As Long
        TmpCutGridX = IndexX '+ HideCountFixCol
        TmpCutGridY = IndexY
        If GridX(TmpCutGridX).GWidthDefault = False Then GridX(TmpCutGridX).GridWidth = SetNewGrid.GridSize.SellWidth_Def
            Cx = GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth
            Cy = GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight
        If GridXYData(TmpCutGridX, TmpCutGridY).Grid.GColorDefault(0) = False Then
            BColor = SetNewGrid.GridXYBackColor
        Else
            BColor = GridXYData(TmpCutGridX, TmpCutGridY).Grid.BackColor
        End If
            If GridXYData(TmpCutGridX, TmpCutGridY).Grid.GColorDefault(0) = False Then
                Picture1.Line (GridX(TmpCutGridX).GridLeft + 2, GridY(TmpCutGridY).GridTop + 0)- _
                (Cx - 2, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.SellHeight_Def), BColor, BF '  &H8000000F, BF
            End If
        'Picture1.Cls
    '    DoEvents
        DrawDrid Picture1, TmpCutGridX, TmpCutGridY, 0, 1
    '    Picture1.Picture = Picture1.Image
    End If
End Property
Sub SellTextSetNoHit(IndexX As Long, IndexY As Long, Texts As String)
    GridXYData(IndexX, IndexY).GridXYValue = Texts
End Sub

Property Get SellSubText(ByVal IndexX As Long, ByVal IndexY As Long) As String
    SellSubText = GridXYData(IndexX, IndexY).GridXYValueSub
End Property
Property Let SellSubText(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As String)
    GridXYData(IndexX, IndexY).GridXYValueSub = NewValue
    If GridX(IndexX).Visibles = True Then Exit Property
    If GridX(IndexX).GridLeft <= 0 Then Exit Property
    
'    If ShowFixCol(0) <> IndexX And GridX(IndexX).GridRealPosisi <= 0 Then GridX(IndexX).GridRealPosisi = IndexX
'    If (GridX(IndexX).GridIndexHead > 0 And GridX(IndexX).GridIndexHead >= HscHead.Value And GridX(IndexX).GridIndexHead <= ShowOnConHeadSell + HscHead.Value) Or (GridX(IndexX).GridRealPosisi >= Hsc.Value And GridX(IndexX).GridRealPosisi <= GridRightX) And (IndexY >= Vsc.Value And IndexY <= GridDownY) Then
    If ShowFixCol(0) <> IndexX And GridX(IndexX).GridRealPosisi <= 0 Then GridX(IndexX).GridRealPosisi = IndexX
    If ((GridX(IndexX).GridIndexHead > 0 And GridX(IndexX).GridIndexHead > HscHead.Value And GridX(IndexX).GridIndexHead <= ShowOnConHeadSell + HscHead.Value) And (GridX(IndexX).GridLeft >= SetNewGrid.GridSize.GDRangeX And GridX(IndexX).GridLeft <= Picture1.ScaleWidth) And (GridY(IndexY).GridTop >= SetNewGrid.GridSize.GDRangeY And GridY(IndexY).GridTop <= Picture1.ScaleHeight)) _
       Or (GridX(IndexX).GridOnIndexHead = 0 And (GridX(IndexX).GridRealPosisi >= Hsc.Value And GridX(IndexX).GridRealPosisi <= GridRightX) And (IndexY >= Vsc.Value And IndexY <= GridDownY)) Then
        
        Dim TmpCutGridX As Integer, TmpCutGridY As Integer, BColor As Long
        TmpCutGridX = IndexX
        TmpCutGridY = IndexY
        If GridX(TmpCutGridX).GWidthDefault = False Then GridX(TmpCutGridX).GridWidth = SetNewGrid.GridSize.SellWidth_Def
            Cx = GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth
            Cy = GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight
        If GridXYData(TmpCutGridX, TmpCutGridY).Grid.GColorDefault(0) = False Then
            BColor = SetNewGrid.GridXYBackColor
        Else
            BColor = GridXYData(TmpCutGridX, TmpCutGridY).Grid.BackColor
        End If
            'If GridXYData(TmpCutGridX, TmpCutGridY).Grid.GColorDefault(0) = False Then
            '    Picture1.Line (GridX(TmpCutGridX).GridLeft + 2, GridY(TmpCutGridY).GridTop + 0)- _
            '    (Cx - 2, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.SellHeight_Def), BColor, BF '  &H8000000F, BF
            'End If
        DrawDrid Picture1, TmpCutGridX, TmpCutGridY, 0, 2
'        Picture1.Picture = Picture1.Image
    End If
End Property
Sub SellSubTextSetNoHit(IndexX As Long, IndexY As Long, Texts As String)
    GridXYData(IndexX, IndexY).GridXYValueSub = Texts
End Sub

'NEW13
Property Get SellListCC(ByVal IndexX As Long, ByVal IndexY As Long, IndexList As Long) As GList
    Set SellListCC = GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList)
End Property
Property Let SellListCC(ByVal IndexX As Long, ByVal IndexY As Long, IndexList As Long, ByVal NewValue As GList)
    Set GridXYData(IndexX, IndexY).GridSubType.ContGList(IndexList) = NewValue
End Property

Property Get SellListType(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    SellListType = GridXYData(IndexX, IndexY).GridSubType.TypeControl
End Property
Property Let SellListType(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
    GridXYData(IndexX, IndexY).GridSubType.TypeControl = NewValue
End Property

'GridXYData(0, 0).GridSubType.TypeControl = 1
'------------------------------------------------------------------------------------------

Property Get SellAlignment(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    SellAlignment = GridXYData(IndexX, IndexY).Grid.Alignment
End Property
Property Let SellAlignment(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
    If NewValue > 2 Then NewValue = 0
    GridXYData(IndexX, IndexY).Grid.Alignment = NewValue
End Property
Property Get SellSubAlignment(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    SellSubAlignment = GridXYData(IndexX, IndexY).GridSub.Alignment
End Property
Property Let SellSubAlignment(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
    If NewValue > 2 Then NewValue = 0
    GridXYData(IndexX, IndexY).GridSub.Alignment = NewValue
End Property


Property Get SellBold(ByVal IndexX As Long, ByVal IndexY As Long) As Boolean
    SellBold = GridXYData(IndexX, IndexY).Grid.Bold
End Property
Property Let SellBold(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Boolean)
    GridXYData(IndexX, IndexY).Grid.Bold = NewValue
End Property
Property Get SellSubBold(ByVal IndexX As Long, ByVal IndexY As Long) As Boolean
    SellSubBold = GridXYData(IndexX, IndexY).GridSub.Bold
End Property
Property Let SellSubBold(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Boolean)
    GridXYData(IndexX, IndexY).GridSub.Bold = NewValue
End Property


Property Get SellItalic(ByVal IndexX As Long, ByVal IndexY As Long) As Boolean
    SellItalic = GridXYData(IndexX, IndexY).Grid.Italic
End Property
Property Let SellItalic(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Boolean)
    GridXYData(IndexX, IndexY).Grid.Italic = NewValue
End Property
Property Get SellSubItalic(ByVal IndexX As Long, ByVal IndexY As Long) As Boolean
    SellSubItalic = GridXYData(IndexX, IndexY).GridSub.Italic
End Property
Property Let SellSubItalic(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Boolean)
    GridXYData(IndexX, IndexY).GridSub.Italic = NewValue
End Property

Property Get SellUnderline(ByVal IndexX As Long, ByVal IndexY As Long) As Boolean
    SellUnderline = GridXYData(IndexX, IndexY).Grid.Underline
End Property
Property Let SellUnderline(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Boolean)
    GridXYData(IndexX, IndexY).Grid.Underline = NewValue
End Property
Property Get SellSubUnderline(ByVal IndexX As Long, ByVal IndexY As Long) As Boolean
    SellSubUnderline = GridXYData(IndexX, IndexY).GridSub.Underline
End Property
Property Let SellSubUnderline(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Boolean)
    GridXYData(IndexX, IndexY).GridSub.Underline = NewValue
End Property

Property Get SellBackColor_Def() As Long
    SellBackColor_Def = SetNewGrid.GridXYBackColor
End Property
Property Let SellBackColor_Def(ByVal NewValue As Long)
    SetNewGrid.GridXYBackColor = NewValue
End Property
Property Get SellSubBackColor_Def() As Long
    SellSubBackColor_Def = SetNewGrid.GridXYBackColorSub
End Property
Property Let SellSubBackColor_Def(ByVal NewValue As Long)
    SetNewGrid.GridXYBackColorSub = NewValue
End Property

Property Get SellBackColor(ByVal IndexX As Long, ByVal IndexY As Long) As OLE_COLOR
    If GridXYData(IndexX, IndexY).Grid.GColorDefault(0) = False Then
        SellBackColor = SetNewGrid.GridXYBackColor
    Else
        SellBackColor = GridXYData(IndexX, IndexY).Grid.BackColor
    End If
End Property
Property Let SellBackColor(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As OLE_COLOR)
    GridXYData(IndexX, IndexY).Grid.BackColor = NewValue
        GridXYData(IndexX, IndexY).Grid.GColorDefault(0) = True
    
    Exit Property 'Cut-------------
    Dim TmpCutGridX As Integer, TmpCutGridY As Integer ', BColor As Long
    TmpCutGridX = IndexX
    TmpCutGridY = IndexY
    
    DrawDrid Picture1, TmpCutGridX, TmpCutGridY, 0
End Property
Property Get SellSubBackColor(ByVal IndexX As Long, ByVal IndexY As Long) As OLE_COLOR
    SellSubBackColor = GridXYData(IndexX, IndexY).GridSub.BackColor
End Property
Property Let SellSubBackColor(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As OLE_COLOR)
    GridXYData(IndexX, IndexY).GridSub.BackColor = NewValue
        GridXYData(IndexX, IndexY).GridSub.GColorDefault(0) = True
End Property

Property Get SellFillColor_Def() As Long
    SellFillColor_Def = SetNewGrid.GridXYFillColor
End Property
Property Let SellFillColor_Def(ByVal NewValue As Long)
    SetNewGrid.GridXYForeColor = NewValue
End Property
Property Get SellSubFillColor_Def() As Long
    SellSubFillColor_Def = SetNewGrid.GridXYFillColorSub
End Property
Property Let SellSubFillColor_Def(ByVal NewValue As Long)
    SetNewGrid.GridXYFillColorSub = NewValue
End Property

Property Get SellFillColor(ByVal IndexX As Long, ByVal IndexY As Long) As OLE_COLOR
    SellFillColor = GridXYData(IndexX, IndexY).Grid.FillColor
End Property
Property Let SellFillColor(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As OLE_COLOR)
    GridXYData(IndexX, IndexY).Grid.FillColor = NewValue
        GridXYData(IndexX, IndexY).Grid.GColorDefault(1) = True
    
    Exit Property 'Cut-------------
    Dim TmpCutGridX As Integer, TmpCutGridY As Integer ', BColor As Long
    TmpCutGridX = IndexX
    TmpCutGridY = IndexY
    
    DrawDrid Picture1, TmpCutGridX, TmpCutGridY, 0
End Property
Property Get SellSubFillColor(ByVal IndexX As Long, ByVal IndexY As Long) As OLE_COLOR
    SellSubFillColor = GridXYData(IndexX, IndexY).GridSub.FillColor
End Property
Property Let SellSubFillColor(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As OLE_COLOR)
    GridXYData(IndexX, IndexY).GridSub.FillColor = NewValue
        GridXYData(IndexX, IndexY).GridSub.GColorDefault(1) = True
End Property





Property Get SellForeColor_Def() As Long
    SellForeColor_Def = SetNewGrid.GridXYForeColor
End Property
Property Let SellForeColor_Def(ByVal NewValue As Long)
    SetNewGrid.GridXYForeColor = NewValue
End Property
Property Get SellSubForeColor_Def() As Long
    SellSubForeColor_Def = SetNewGrid.GridXYForeColorSub
End Property
Property Let SellSubForeColor_Def(ByVal NewValue As Long)
    SetNewGrid.GridXYForeColorSub = NewValue
End Property


Property Get SellForeColor(ByVal IndexX As Long, ByVal IndexY As Long) As OLE_COLOR
    SellForeColor = GridXYData(IndexX, IndexY).Grid.ForeColor
End Property
Property Let SellForeColor(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As OLE_COLOR)
    GridXYData(IndexX, IndexY).Grid.ForeColor = NewValue
'        GridXYData(IndexX, IndexY).GColorDefault(1) = True

    Exit Property 'Cut-------------
    Dim TmpCutGridX As Integer, TmpCutGridY As Integer ', BColor As Long
    TmpCutGridX = IndexX
    TmpCutGridY = IndexY

    DrawDrid Picture1, TmpCutGridX, TmpCutGridY, 0
End Property
Property Get SellSubForeColor(ByVal IndexX As Long, ByVal IndexY As Long) As OLE_COLOR
    SellSubForeColor = GridXYData(IndexX, IndexY).GridSub.ForeColor
End Property
Property Let SellSubForeColor(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As OLE_COLOR)
    GridXYData(IndexX, IndexY).GridSub.ForeColor = NewValue
End Property
'---------------------------------------------------------------------------------------------------------------------------------
'***********************************************************************************************************************************

Property Get SellIconX1() As Integer
    SellIconX1 = SetNewGrid.GridSizePic.SellIconX1
End Property
Property Let SellIconX1(ByVal NewValue As Integer)
    SetNewGrid.GridSizePic.SellIconX1 = NewValue
End Property
Property Get SellIconY1() As Integer
    SellIconY1 = SetNewGrid.GridSizePic.SellIconY1
End Property
Property Let SellIconY1(ByVal NewValue As Integer)
    SetNewGrid.GridSizePic.SellIconY1 = NewValue
End Property

Property Get SellIconX2() As Integer
    SellIconX2 = SetNewGrid.GridSizePic.SellIconX2
End Property
Property Let SellIconX2(ByVal NewValue As Integer)
    SetNewGrid.GridSizePic.SellIconX2 = NewValue
End Property
Property Get SellIconY2() As Integer
    SellIconY2 = SetNewGrid.GridSizePic.SellIconY2
End Property
Property Let SellIconY2(ByVal NewValue As Integer)
    SetNewGrid.GridSizePic.SellIconY2 = NewValue
End Property


Property Get PictureErr() As Boolean 'Belom ada coment
    PictureErr = PicErr
End Property
Property Let PictureErr(ByVal NewValue As Boolean)
    PicErr = NewValue
End Property

Property Get SellIconFilePicture() As String 'Belom ada coment
    SellIconFilePicture = SetNewGrid.GridFilePicture
End Property
Property Let SellIconFilePicture(ByVal NewValue As String)
    SetNewGrid.GridFilePicture = NewValue
End Property

Property Get SellIconIndex(ByVal IndexX As Long, ByVal IndexY As Long) As Integer
    SellIconIndex = GridXYData(IndexX, IndexY).GridXYPicIndex
End Property
Property Let SellIconIndex(ByVal IndexX As Long, ByVal IndexY As Long, ByVal NewValue As Integer)
Dim IndexX1 As Integer, IndexY1 As Integer
'Dim TmpCutGridX As Integer, TmpCutGridY As Integer, BColor As Long
    
    GridXYData(IndexX, IndexY).GridXYPicIndex = NewValue
    If GridX(IndexX).Visibles = True Then Exit Property
    If GridX(IndexX).GridLeft <= 0 Then Exit Property
    
    If NewValue > -1 And ((GridX(IndexX).GridIndexHead > 0 And GridX(IndexX).GridIndexHead <= ShowOnConHeadSell) Or (IndexX >= Hsc.Value And IndexX <= GridRightX) And (IndexY >= Vsc.Value And IndexY <= GridDownY)) Then
        IndexX1 = IndexX: IndexY1 = IndexY
        DrawPicGrid Picture1, NewValue, IndexX1, IndexY1
    End If
End Property

Property Get SellIconPicContColms() As Integer 'Belom ada coment
    SellIconPicContColms = SetNewGrid.GridSizePic.SellIconPicContColms
End Property
Property Let SellIconPicContColms(ByVal NewValue As Integer)
    SetNewGrid.GridSizePic.SellIconPicContColms = NewValue
End Property
Property Get SellIconPicContRows() As Integer 'Belom ada coment
    SellIconPicContRows = SetNewGrid.GridSizePic.SellIconPicContRows
End Property
Property Let SellIconPicContRows(ByVal NewValue As Integer)
    SetNewGrid.GridSizePic.SellIconPicContRows = NewValue
End Property

Property Get SellIcon(ByVal Index As Long) As Boolean 'Belom ada coment

    If ShowOnConHeadSell > 0 And Index < 0 Then
        Index = HeadSellOnFixCol(Abs(Index) - 1)
    End If
    SellIcon = GridX(Index).GPicturePut
End Property
Property Let SellIcon(ByVal Index As Long, ByVal NewValue As Boolean)
    GridX(Index).GPicturePut = NewValue
End Property








































Property Get GridType() As Integer 'Belom ada coment
    GridType = SetNewGrid.GridType
End Property
Property Let GridType(ByVal NewValue As Integer)
    SetNewGrid.GridType = NewValue
End Property








Property Get RangePicSubX1() As Integer 'Belom ada coment
    RangePicSubX1 = SetNewGrid.GridSizePicSub.RangePicSubX1
End Property
Property Let RangePicSubX1(ByVal NewValue As Integer)
    SetNewGrid.GridSizePicSub.RangePicSubX1 = NewValue
End Property

Property Get RangePicSubY1() As Integer 'Belom ada coment
    RangePicSubY1 = SetNewGrid.GridSizePicSub.RangePicSubY1
End Property
Property Let RangePicSubY1(ByVal NewValue As Integer)
    SetNewGrid.GridSizePicSub.RangePicSubY1 = NewValue
End Property

Property Get RangePicSubX2() As Integer 'Belom ada coment
    RangePicSubX2 = SetNewGrid.GridSizePicSub.RangePicSubX2
End Property
Property Let RangePicSubX2(ByVal NewValue As Integer)
    SetNewGrid.GridSizePicSub.RangePicSubX2 = NewValue
End Property

Property Get RangePicSubY2() As Integer 'Belom ada coment
    RangePicSubY2 = SetNewGrid.GridSizePicSub.RangePicSubY2
End Property
Property Let RangePicSubY2(ByVal NewValue As Integer)
    SetNewGrid.GridSizePicSub.RangePicSubY2 = NewValue
End Property


Property Get TableWidth() As Single 'Belom ada coment
    TableWidth = SetNewGrid.GridSize.TableWidth
End Property
Property Let TableWidth(ByVal NewValue As Single)
    SetNewGrid.GridSize.TableWidth = NewValue
End Property

Property Get TableHeight() As Single 'Belom ada coment
    TableHeight = SetNewGrid.GridSize.TableHeight
End Property
Property Let TableHeight(ByVal NewValue As Single)
    SetNewGrid.GridSize.TableHeight = NewValue
End Property

Property Get NameTabel() As String 'Belom ada coment
    NameTabel = SetNewGrid.GridXYName
End Property
Property Let NameTabel(ByVal NewValue As String)
    SetNewGrid.GridXYName = NewValue
End Property







'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


'-------------------------------------------------------------------------------






Property Get SellNoFill(ByVal IndexX As Long, ByVal IndexY As Long, ByVal Nomber As Integer) As Boolean
    SellNoFill = GridXYData(IndexX, IndexY).Grid.GColorDefault(Nomber)
End Property
Property Let SellNoFill(ByVal IndexX As Long, ByVal IndexY As Long, ByVal Nomber As Integer, ByVal NewValue As Boolean)
    GridXYData(IndexX, IndexY).Grid.GColorDefault(Nomber) = NewValue
End Property

Property Get SellNoFillSub(ByVal IndexX As Long, ByVal IndexY As Long, ByVal Nomber As Integer) As Boolean
    SellNoFillSub = GridXYData(IndexX, IndexY).GridSub.GColorDefault(Nomber)
End Property
Property Let SellNoFillSub(ByVal IndexX As Long, ByVal IndexY As Long, ByVal Nomber As Integer, ByVal NewValue As Boolean)
    GridXYData(IndexX, IndexY).GridSub.GColorDefault(Nomber) = NewValue
End Property



'bx91

'Batas





'Sub SizeLeftRight(Index As Long, NewValue As Boolean)
'End Sub





'-------------------------------------------------------------------------- y
'FixdRowsColStyl









'Batas +++++++++++++++++++++++++++++++++++++++

Sub SizeUpDown(Index As Long, NewValue As Boolean)
    If NewValue = False Then
        GridY(Index).GHSave = Abs(GridY(Index).GridHeight - SetNewGrid.GridSize.SellHeight_Def)
        GridY(Index).GHeightDefault = False
    Else
        If GridY(Index).GHSave = 0 Or GridY(Index).GHSave = SetNewGrid.GridSize.SellHeight_Def Then GridY(Index).GHSave = 80  'Add
        GridY(Index).GridHeight = GridY(Index).GHSave
        GridY(Index).GHeightDefault = True
    End If
End Sub

Sub TMP_GridXY(ObjMe As Object, GridLeftX As Integer, GridUpY As Integer, Optional ShowPic As Boolean, Optional ShowX As Long = -1, Optional ShowY As Long = -1, Optional NoDrawing As Boolean)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim TFontWidth As Long, TMPConHeadSellOnFixCol As Integer
Dim ProsesShow As Integer, ProsesShowMin As Integer
Dim IndexHeadEnd As Integer, ShowConHeadSellOnFixCol As Integer
Dim TmpGridRightX As Integer
Dim TmpColorsLineHead As Long, CloseHeadFix As Boolean
Dim RangeEndHead As Integer, LineBack As Integer

TmpCutGridX = GridLeftX
TmpCutGridY = GridUpY

GridX(ShowFixCol(TmpCutGridX)).GridLeft = SetNewGrid.GridSize.GDRangeX
GridY(GridUpY).GridTop = SetNewGrid.GridSize.GDRangeY

ObjMe.Line (SetNewGrid.GridSize.GDRangeX, SetNewGrid.GridSize.GDRangeY)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), SetNewGrid.GridXYBackColor, BF

If SetNewGrid.GridStyleX = 0 Then
    ObjMe.Line (SetNewGrid.GridSize.GDRangeX, 0)- _
               (ObjMe.ScaleWidth, SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleX = 1 Or SetNewGrid.GridStyleX = 2 Then
    pvDrawGrid ObjMe, Val(SetNewGrid.GridSize.GDRangeX), 0, ObjMe.ScaleWidth, Val(SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), SetNewGrid.GridLenkapX.GridBackColGra(1), SetNewGrid.GridStyleX - 1
End If
If SetNewGrid.GridStyleY = 0 Then
    ObjMe.Line (0, SetNewGrid.GridSize.GDRangeY)- _
               (SetNewGrid.GridSize.GDRangeX, ObjMe.ScaleHeight), SetNewGrid.GridLenkapY.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleY = 1 Or SetNewGrid.GridStyleY = 2 Then
    pvDrawGrid ObjMe, 0, Val(SetNewGrid.GridSize.GDRangeY), Val(SetNewGrid.GridSize.GDRangeX), ObjMe.ScaleHeight, SetNewGrid.GridLenkapY.GridBackColGra(0), SetNewGrid.GridLenkapY.GridBackColGra(1), SetNewGrid.GridStyleY - 1
End If

''On Error GoTo papa
''ConHeadSellOnFixCol = -1
''If 8 = 8 Then
''    Static Vion As Integer
    
''    ConHeadSellOnFixCol = 3
''    ReDim Preserve HeadSellOnFixCol(ConHeadSellOnFixCol - 1)
''    HeadSellOnFixCol(0) = 0 '.IndexFixCol = 0
'''    GridX(HeadSellOnFixCol(0)).GridRealOnPosisi
'''    HeadSellOnFixCol(0).RealPosision = 0

''    HeadSellOnFixCol(1) = 2
''    HeadSellOnFixCol(2) = 10
'''    HeadSellOnFixCol(3) = 25
'''    HeadSellOnFixCol(4) = 28
'''    HeadSellOnFixCol(5) = 30
'''    HeadSellOnFixCol(6) = 35
''End If
'papax:
HeadCountAuto = False
RangeEndHead = 300 '>>>>>>>>>>>>>>>>

Command25.Caption = ""
TMPConHeadSellOnFixCol = 0 ' Vion
IndexHeadEnd = -1

ShowOnConHeadSell = 0
'GridX(*** ShowFixCol(TmpCutGridX)).GridLeft = 150
TMPConHeadSellOnFixCol = HscHead.Value
Do 'X
    ProsesShow = ShowFixCol(TmpCutGridX)
    If TmpCutGridX > GridLeftX Then ProsesShowMin = ShowFixCol(TmpCutGridX - 1)
    'If ShowConHeadSellOnFixCol > 2 Then CloseHeadFix = True
    If 8 = 8 Then
    If IndexHeadEnd <> -1 Then
        ProsesShowMin = IndexHeadEnd 'HeadSellOnFixCol(Vion)
        IndexHeadEnd = -1
    End If
    GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = 0

    If TMPConHeadSellOnFixCol < ConHeadSellOnFixCol Then
        If GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = -1 Then _
            GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        If GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = -1 Then _
            GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = ShowFixCol(TmpCutGridX)

        'If TMPConHeadSellOnFixCol > 2 + HscHead.Value Then CloseHeadFix = True
        If CloseHeadFix = False And GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi >= GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi Then
            If ShowFixCol(TmpCutGridX) <> HeadSellOnFixCol(TMPConHeadSellOnFixCol) Then
                ProsesShow = HeadSellOnFixCol(TMPConHeadSellOnFixCol)   '= 1
                GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridLeft = SetNewGrid.GridSize.GDRangeX  '+ 100
                
                If TmpCutGridX > GridLeftX Then ProsesShowMin = HeadSellOnFixCol(TMPConHeadSellOnFixCol - 1)
                ShowConHeadSellOnFixCol = ShowConHeadSellOnFixCol + 1
                ShowOnConHeadSell = ShowConHeadSellOnFixCol
                   
                GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = TMPConHeadSellOnFixCol + 1
                Command25.Caption = Command25.Caption & TmpCutGridX & " " & TMPConHeadSellOnFixCol & "-" & HeadSellOnFixCol(TMPConHeadSellOnFixCol) & ". "
            Else
            End If
            IndexHeadEnd = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        Else
            TMPConHeadSellOnFixCol = ConHeadSellOnFixCol
        End If
    Else
    End If
    
    TMPConHeadSellOnFixCol = TMPConHeadSellOnFixCol + 1
    End If
    '--------------
    If ShowConHeadSellOnFixCol > 0 And TmpCutGridX = GridLeftX + ShowConHeadSellOnFixCol Then
        RangeEndHead = 5
        'RangeEndHead = 250 '>>>>>>>>>>>>>>>>
    Else
        RangeEndHead = 0
        'RangeEndHead = 0
    End If
    If GridX(ProsesShow).GWidthDefault = False Then GridX(ProsesShow).GridWidth = SetNewGrid.GridSize.SellWidth_Def  '100
'    If HeadCountAuto = False Then
'MsgBox ProsesShow & " " & ProsesShowMin
    If TmpCutGridX > GridLeftX Then _
    GridX(ProsesShow).GridLeft = GridX(ProsesShowMin).GridLeft + GridX(ProsesShowMin).GridWidth + RangeEndHead '- RangeEndHead
    
'    If ShowConHeadSellOnFixCol = 3 And GridLeftX + ShowConHeadSellOnFixCol = ProsesShow Then
    If GridX(ProsesShow).GridRealOnPosisi = -1 Then GridX(ProsesShow).GridRealOnPosisi = ShowFixCol(ProsesShow)
    If ShowConHeadSellOnFixCol > 0 And GridLeftX + ShowConHeadSellOnFixCol = GridX(ProsesShow).GridRealOnPosisi Then
        ObjMe.Line (RangeEndHead, 0)-(ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF
        GridX(ProsesShow).GridLeft = RangeEndHead + RangeEndHead
    End If

    If GridX(ProsesShow).GridLeft + GridX(ProsesShow).GridWidth >= RangeEndHead Then 'Menutup proses head
'        MsgBox ProsesShow
        CloseHeadFix = True
    End If
    
    Do 'Y
        If GridXYData(ProsesShow, TmpCutGridY).GridSub.GColorDefault(0) = False Then _
        GridXYData(ProsesShow, TmpCutGridY).GridSub.BackColor = SetNewGrid.GridXYBackColorSub
        If GridY(TmpCutGridY).GHeightDefault = False Then
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def
        Else
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def + GridY(TmpCutGridY).GHSave '20
        End If
        If TmpCutGridY > GridUpY Then _
        GridY(TmpCutGridY).GridTop = GridY(TmpCutGridY - 1).GridTop + GridY(TmpCutGridY - 1).GridHeight + 0
        
        If DoitY = False Then DrawDrid ObjMe, ProsesShow, TmpCutGridY, 1, , NoDrawing
        DrawDrid ObjMe, ProsesShow, TmpCutGridY, 0, , NoDrawing
        
        'UNTUK PENAMBAHAN CONTROL PADA SUB GRID
        If GridY(TmpCutGridY).GHeightDefault = True Then
            Select Case GridXYData(ProsesShow, TmpCutGridY).GridSubType.TypeControl '= "1" Then
            Case 1
                DrawList ProsesShow, TmpCutGridY, ObjMe ', Fg, Ad
'            Case 2
'                DrawButton *** ShowFixCol(TmpCutGridX), TmpCutGridY, ObjMe
            End Select
        End If
                
        If GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.TableHeight Or _
        TmpCutGridY + 1 = SetNewGrid.GridYCount Then
            GridDownY = TmpCutGridY '- 10
            TmpCutGridY = GridUpY

            Exit Do
        End If
    TmpCutGridY = TmpCutGridY + 1
    Loop

    DrawDrid ObjMe, ProsesShow, TmpCutGridY, 2, , NoDrawing
    
    If GridX(ProsesShow).GridLeft + GridX(ProsesShow).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    TmpCutGridX = SetNewGrid.GridXCount - HideCountFixCol - 1 Then
        GridRightX = TmpCutGridX  '*** ShowFixCol(TmpCutGridX) '- 10
        
        Exit Do
    End If
TmpCutGridX = TmpCutGridX + 1
DoitY = True
Loop

If ShowOnConHeadSell > 0 Then
    HeadWidFixCol = GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth
'    MsgBox GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth
Else
    HeadWidFixCol = 0
End If

ObjMe.Line (0, GridY(GridDownY).GridTop + GridY(GridDownY).GridHeight + 1)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

'>>>>>>>>>>>>> this are in eror &HC0FFC0
'If SellCountColumn - 1 = GridRightX And ShowOnConHeadSell - 1 = GridRightX - GridLeftX Then
'    TmpGridRightX = HeadSellOnFixCol(ShowOnConHeadSell - 1)
'    TmpColorsLineHead = vbRed
'    If HeadCountAuto = False Then LineBack = GridX(ShowFixCol(TmpGridRightX)).GridLeft + GridX(ShowFixCol(TmpGridRightX)).GridWidth - RangeEndHead
'Else 'HeadCountAuto
'    TmpGridRightX = GridRightX
'    If SellCountColumn - 1 = GridRightX Then TmpColorsLineHead = vbYellow Else TmpColorsLineHead = &HC0FFC0
'End If

ObjMe.Line (GridX(ShowFixCol(TmpGridRightX)).GridLeft + GridX(ShowFixCol(TmpGridRightX)).GridWidth + 1 - LineBack, 0)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

TmpColorsLineHead = LinePosisionHeads
If ShowOnConHeadSell > 0 Then
    If HeadCountAuto = False Then
        DrawingLineHeadCol ObjMe, _
        RangeEndHead, 5, TmpColorsLineHead
    Else
        DrawingLineHeadCol ObjMe, _
        GridX(HeadSellOnFixCol(ShowConHeadSellOnFixCol - 1 + HscHead.Value)).GridLeft + GridX(HeadSellOnFixCol(ShowConHeadSellOnFixCol - 1 + HscHead.Value)).GridWidth, 5, TmpColorsLineHead
    End If
End If

If Text1.Visible = True Then _
AllText Text1, SellText(FixdColnIndex(XIndexPointerText1), YIndexPointerText1), SellLeft(FixdColnIndex(XIndexPointerText1)), SellTop(YIndexPointerText1), SellWidth(FixdColnIndex(XIndexPointerText1)), SellHeight_Def
''On AllText Text1, SellText(XIndexPointerText1, YIndexPointerText1), SellLeft(XIndexPointerText1), SellTop(YIndexPointerText1), SellWidth(XIndexPointerText1), SellHeight_Def

'HscHead.Value >>>>>>>

Command24.Caption = ShowOnConHeadSell & " " & GridLeftX & " " & GridRightX - GridLeftX

'Command23.Caption = " "
'For iX = 0 To 9
'    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridLeft & ", "
'Next iX

'Exit Sub
Command24.Caption = ShowConHeadSellOnFixCol & "."
On Error GoTo papa
Command22.Caption = "HeadSellOnFixCol "
For iX = 0 To ConHeadSellOnFixCol - 1
    Command22.Caption = Command22.Caption & GridX(HeadSellOnFixCol(iX)).GridIndexHead - 1 & "." & HeadSellOnFixCol(iX) & " "
Next iX

Command23.Caption = "GridRealOnPosisi "
For iX = 0 To 12
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealOnPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridRealPosisi "
For iX = 0 To 12
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "ShowFixCol [" & UBound(ShowFixCol()) & "] "
For iX = 0 To 12 'UBound(ShowFixCol())
    Command23.Caption = Command23.Caption & iX & "." & ShowFixCol(iX) & " "  '& "." & GridX(ShowFixCol(iX)).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridIndexHead "
For iX = 0 To 12 'UBound(ShowFixCol())
    If GridX(ShowFixCol(iX)).GridIndexHead <> 0 Then
    Command23.Caption = Command23.Caption & iX & "." & ShowFixCol(iX) & ">" & GridX(ShowFixCol(iX)).GridIndexHead - 1 & ", " '[" & iX & ". " & GridX(iX).GridIndexHead & "] "
    End If
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridOnIndexHead "
For iX = 0 To 12 'UBound(ShowFixCol())
    'If GridX(ShowFixCol(iX)).GridIndexHead <> 0 Then
    Command23.Caption = Command23.Caption & iX & "." & GridX(ShowFixCol(iX)).GridOnIndexHead & ", "
    'End If UBound(ShowFixCol())
Next iX

Command29.Caption = ""
For iX = 0 To 5 'UBound(ShowFixCol())
    Command29.Caption = Command29.Caption & iX & "." & HeadSellOnFixCol(iX) & " " '& ">" & GridX(ShowFixCol(iX)).GridIndexHead - 1 & ", " '[" & iX & ". " & GridX(iX).GridIndexHead & "] "
Next iX
papa:
'Command22.Caption = "000000"dad
End Sub

Sub GridXY(ObjMe As Object, GridLeftX As Integer, GridUpY As Integer, Optional ShowPic As Boolean, Optional ShowX As Long = -1, Optional ShowY As Long = -1, Optional NoDrawing As Boolean)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim TFontWidth As Long, TMPConHeadSellOnFixCol As Integer
Dim ProsesShow As Integer, ProsesShowMin As Integer
Dim IndexHeadEnd As Integer, ShowConHeadSellOnFixCol As Integer
Dim TmpGridRightX As Integer
Dim TmpColorsLineHead As Long, CloseHeadFix As Boolean
Dim RangeEndHead_Range As Integer, LineBack As Integer

TmpCutGridX = GridLeftX
TmpCutGridY = GridUpY

GridX(ShowFixCol(TmpCutGridX)).GridLeft = SetNewGrid.GridSize.GDRangeX
GridY(GridUpY).GridTop = SetNewGrid.GridSize.GDRangeY

ObjMe.Line (SetNewGrid.GridSize.GDRangeX, SetNewGrid.GridSize.GDRangeY)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), SetNewGrid.GridXYBackColor, BF

If SetNewGrid.GridStyleX = 0 Then
    ObjMe.Line (SetNewGrid.GridSize.GDRangeX, 0)- _
               (ObjMe.ScaleWidth, SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleX = 1 Or SetNewGrid.GridStyleX = 2 Then
    pvDrawGrid ObjMe, Val(SetNewGrid.GridSize.GDRangeX), 0, ObjMe.ScaleWidth, Val(SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), SetNewGrid.GridLenkapX.GridBackColGra(1), SetNewGrid.GridStyleX - 1
End If
If SetNewGrid.GridStyleY = 0 Then
    ObjMe.Line (0, SetNewGrid.GridSize.GDRangeY)- _
               (SetNewGrid.GridSize.GDRangeX, ObjMe.ScaleHeight), SetNewGrid.GridLenkapY.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleY = 1 Or SetNewGrid.GridStyleY = 2 Then
    pvDrawGrid ObjMe, 0, Val(SetNewGrid.GridSize.GDRangeY), Val(SetNewGrid.GridSize.GDRangeX), ObjMe.ScaleHeight, SetNewGrid.GridLenkapY.GridBackColGra(0), SetNewGrid.GridLenkapY.GridBackColGra(1), SetNewGrid.GridStyleY - 1
End If

''On Error GoTo papa
''ConHeadSellOnFixCol = -1
''If 8 = 8 Then
''    Static Vion As Integer
    
''    ConHeadSellOnFixCol = 3
''    ReDim Preserve HeadSellOnFixCol(ConHeadSellOnFixCol - 1)
''    HeadSellOnFixCol(0) = 0 '.IndexFixCol = 0
'''    GridX(HeadSellOnFixCol(0)).GridRealOnPosisi
'''    HeadSellOnFixCol(0).RealPosision = 0

''    HeadSellOnFixCol(1) = 2
''    HeadSellOnFixCol(2) = 10
'''    HeadSellOnFixCol(3) = 25
'''    HeadSellOnFixCol(4) = 28
'''    HeadSellOnFixCol(5) = 30
'''    HeadSellOnFixCol(6) = 35
''End If
'papax:

''HeadCountAuto = True '>>>>>>>>>>>>>>>>>>> di perbaiki
''RangeEndHead = 300 '>>>>>>>>>>>>>>>>
''Dim RangeEndHead_PosAuto As Integer
''RangeEndHead_PosAuto = 3


Command25.Caption = ""
TMPConHeadSellOnFixCol = 0 ' Vion
IndexHeadEnd = -1

ShowOnConHeadSell = 0
'GridX(*** ShowFixCol(TmpCutGridX)).GridLeft = 150
TMPConHeadSellOnFixCol = HscHead.Value
Do 'X
    ProsesShow = ShowFixCol(TmpCutGridX)
    If TmpCutGridX > GridLeftX Then ProsesShowMin = ShowFixCol(TmpCutGridX - 1)
    'If ShowConHeadSellOnFixCol > 2 Then CloseHeadFix = True
    If 8 = 8 Then
    If IndexHeadEnd <> -1 Then
        ProsesShowMin = IndexHeadEnd 'HeadSellOnFixCol(Vion)
        IndexHeadEnd = -1
    End If
    GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = 0

    If TMPConHeadSellOnFixCol < ConHeadSellOnFixCol Then
        If GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = -1 Then _
            GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        If GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = -1 Then _
            GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = ShowFixCol(TmpCutGridX)

        'If TMPConHeadSellOnFixCol > 2 + HscHead.Value Then CloseHeadFix = True
        If CloseHeadFix = False And GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi >= GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi Then
            If ShowFixCol(TmpCutGridX) <> HeadSellOnFixCol(TMPConHeadSellOnFixCol) Then
                ProsesShow = HeadSellOnFixCol(TMPConHeadSellOnFixCol)   '= 1
                GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridLeft = SetNewGrid.GridSize.GDRangeX  '+ 100
                
                If TmpCutGridX > GridLeftX Then ProsesShowMin = HeadSellOnFixCol(TMPConHeadSellOnFixCol - 1)
                ShowConHeadSellOnFixCol = ShowConHeadSellOnFixCol + 1
                ShowOnConHeadSell = ShowConHeadSellOnFixCol
                   
                GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = TMPConHeadSellOnFixCol + 1
                Command25.Caption = Command25.Caption & TmpCutGridX & " " & TMPConHeadSellOnFixCol & "-" & HeadSellOnFixCol(TMPConHeadSellOnFixCol) & ". "
            Else
            End If
            IndexHeadEnd = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        Else
            TMPConHeadSellOnFixCol = ConHeadSellOnFixCol
        End If
    Else
    End If
    
    TMPConHeadSellOnFixCol = TMPConHeadSellOnFixCol + 1
    End If
    '--------------
    If ShowConHeadSellOnFixCol > 0 And TmpCutGridX = GridLeftX + ShowConHeadSellOnFixCol Then
        RangeEndHead_Range = 5
    Else
        RangeEndHead_Range = 0
    End If
    If GridX(ProsesShow).GWidthDefault = False Then GridX(ProsesShow).GridWidth = SetNewGrid.GridSize.SellWidth_Def  '100
'    If HeadCountAuto = False Then
'MsgBox ProsesShow & " " & ProsesShowMin
    If TmpCutGridX > GridLeftX Then _
    GridX(ProsesShow).GridLeft = GridX(ProsesShowMin).GridLeft + GridX(ProsesShowMin).GridWidth + RangeEndHead_Range '- RangeEndHead
    
'    If ShowConHeadSellOnFixCol = 3 And GridLeftX + ShowConHeadSellOnFixCol = ProsesShow Then
    If GridX(ProsesShow).GridRealOnPosisi = -1 Then GridX(ProsesShow).GridRealOnPosisi = ShowFixCol(ProsesShow)
'' Jika Ya    If GridX(ProsesShow).GridRealPosisi = 0 Then GridX(ProsesShow).GridRealPosisi = ShowFixCol(ProsesShow)

''    If ShowConHeadSellOnFixCol > 0 And GridLeftX + ShowConHeadSellOnFixCol = GridX(ProsesShow).GridRealOnPosisi Then
''Original    If ShowConHeadSellOnFixCol > 0 And GridLeftX + ShowConHeadSellOnFixCol = GridX((ProsesShow)).GridRealPosisi - 0 Then  '...-> Error
    If ShowConHeadSellOnFixCol > 0 And GridLeftX + ShowConHeadSellOnFixCol = FixdColnIndex_Real(ProsesShow) Then  '...-> Error
        If HeadCountAuto = False Then
            ObjMe.Line (RangeEndHead, 0)-(ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF
            GridX(ProsesShow).GridLeft = RangeEndHead + RangeEndHead_Range
        Else
            RangeEndHead = GridX(ProsesShowMin).GridLeft + GridX(ProsesShowMin).GridWidth
        End If
        IndexEndHead = ProsesShow
    End If

    If HeadCountAuto = False And GridX(ProsesShow).GridLeft + GridX(ProsesShow).GridWidth >= RangeEndHead Then  'Menutup proses head
        CloseHeadFix = True
    Else
        If ShowOnConHeadSell >= RangeEndHead_PosAuto Then CloseHeadFix = True
    End If
    
    Do 'Y
        If GridXYData(ProsesShow, TmpCutGridY).GridSub.GColorDefault(0) = False Then _
        GridXYData(ProsesShow, TmpCutGridY).GridSub.BackColor = SetNewGrid.GridXYBackColorSub
        If GridY(TmpCutGridY).GHeightDefault = False Then
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def
        Else
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def + GridY(TmpCutGridY).GHSave '20
        End If
        If TmpCutGridY > GridUpY Then _
        GridY(TmpCutGridY).GridTop = GridY(TmpCutGridY - 1).GridTop + GridY(TmpCutGridY - 1).GridHeight + 0
        
        If DoitY = False Then DrawDrid ObjMe, ProsesShow, TmpCutGridY, 1, , NoDrawing, True
        DrawDrid ObjMe, ProsesShow, TmpCutGridY, 0, , NoDrawing, True
        
        'UNTUK PENAMBAHAN CONTROL PADA SUB GRID
        If GridY(TmpCutGridY).GHeightDefault = True Then
            Select Case GridXYData(ProsesShow, TmpCutGridY).GridSubType.TypeControl '= "1" Then
            Case 1
                DrawList ProsesShow, TmpCutGridY, ObjMe ', Fg, Ad
'            Case 2
'                DrawButton *** ShowFixCol(TmpCutGridX), TmpCutGridY, ObjMe
            End Select
        End If
                
        If GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.TableHeight Or _
        TmpCutGridY + 1 = SetNewGrid.GridYCount Then
            GridDownY = TmpCutGridY '- 10
            TmpCutGridY = GridUpY

            Exit Do
        End If
    TmpCutGridY = TmpCutGridY + 1
    Loop

    DrawDrid ObjMe, ProsesShow, TmpCutGridY, 2, , NoDrawing, True
    
    If GridX(ProsesShow).GridLeft + GridX(ProsesShow).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    TmpCutGridX = SetNewGrid.GridXCount - HideCountFixCol - 1 Then
        GridRightX = TmpCutGridX  '*** ShowFixCol(TmpCutGridX) '- 10
        
        Exit Do
    End If
TmpCutGridX = TmpCutGridX + 1
DoitY = True
Loop

If ShowOnConHeadSell > 0 Then
    HeadWidFixCol = GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth
'    MsgBox GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth
Else
    HeadWidFixCol = 0
End If

ObjMe.Line (0, GridY(GridDownY).GridTop + GridY(GridDownY).GridHeight + 1)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

'>>>>>>>>>>>>> this are in eror &HC0FFC0
If 8 = 9 Then
''If SellCountColumn - 1 = GridRightX And ShowOnConHeadSell - 1 = GridRightX - GridLeftX Then
''    TmpGridRightX = HeadSellOnFixCol(ShowOnConHeadSell - 1)
''    TmpColorsLineHead = vbRed
''    If HeadCountAuto = False Then LineBack = GridX(ShowFixCol(TmpGridRightX)).GridLeft + GridX(ShowFixCol(TmpGridRightX)).GridWidth - RangeEndHead
''Else 'HeadCountAuto
''    TmpGridRightX = GridRightX
''    If SellCountColumn - 1 = GridRightX Then TmpColorsLineHead = vbYellow Else TmpColorsLineHead = &HC0FFC0
''End If
End If
'Kesatuan -------------------------
TmpColorsLineHead = LinePosisionHeads(TmpGridRightX)

ObjMe.Line (GridX(ShowFixCol(TmpGridRightX)).GridLeft + GridX(ShowFixCol(TmpGridRightX)).GridWidth + 1 - LineBack, 0)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

If ShowOnConHeadSell > 0 Then
    If HeadCountAuto = False Then
        DrawingLineHeadCol ObjMe, _
        RangeEndHead, 5, TmpColorsLineHead
    Else
        DrawingLineHeadCol ObjMe, _
        GridX(HeadSellOnFixCol(ShowConHeadSellOnFixCol - 1 + HscHead.Value)).GridLeft + GridX(HeadSellOnFixCol(ShowConHeadSellOnFixCol - 1 + HscHead.Value)).GridWidth, 5, TmpColorsLineHead
    End If
End If
'----------------------------------

If Text1.Visible = True Then _
AllText Text1, SellText(FixdColnIndex(XIndexPointerText1), YIndexPointerText1), SellLeft(FixdColnIndex(XIndexPointerText1)), SellTop(YIndexPointerText1), SellWidth(FixdColnIndex(XIndexPointerText1)), SellHeight_Def
''On AllText Text1, SellText(XIndexPointerText1, YIndexPointerText1), SellLeft(XIndexPointerText1), SellTop(YIndexPointerText1), SellWidth(XIndexPointerText1), SellHeight_Def

'HscHead.Value >>>>>>>

Command24.Caption = ShowOnConHeadSell & " " & GridLeftX & " " & GridRightX - GridLeftX

'Command23.Caption = " "
'For iX = 0 To 9
'    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridLeft & ", "
'Next iX

'Exit Sub
Command24.Caption = ShowConHeadSellOnFixCol & "."
On Error GoTo papa
Command22.Caption = "HeadSellOnFixCol "
For iX = 0 To ConHeadSellOnFixCol - 1
    Command22.Caption = Command22.Caption & GridX(HeadSellOnFixCol(iX)).GridIndexHead - 1 & "." & HeadSellOnFixCol(iX) & " "
Next iX

Command23.Caption = "GridRealOnPosisi "
For iX = 0 To 12
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealOnPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridRealPosisi "
For iX = 0 To 12
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "ShowFixCol [" & UBound(ShowFixCol()) & "] "
For iX = 0 To 12 'UBound(ShowFixCol())
    Command23.Caption = Command23.Caption & iX & "." & ShowFixCol(iX) & " "  '& "." & GridX(ShowFixCol(iX)).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridIndexHead "
For iX = 0 To 12 'UBound(ShowFixCol())
    If GridX(ShowFixCol(iX)).GridIndexHead <> 0 Then
    Command23.Caption = Command23.Caption & iX & "." & ShowFixCol(iX) & ">" & GridX(ShowFixCol(iX)).GridIndexHead - 1 & ", " '[" & iX & ". " & GridX(iX).GridIndexHead & "] "
    End If
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridOnIndexHead "
For iX = 0 To 12 'UBound(ShowFixCol())
    'If GridX(ShowFixCol(iX)).GridIndexHead <> 0 Then
    Command23.Caption = Command23.Caption & iX & "." & GridX(ShowFixCol(iX)).GridOnIndexHead & ", "
    'End If UBound(ShowFixCol())
Next iX

Command29.Caption = ""
For iX = 0 To 5 'UBound(ShowFixCol())
    Command29.Caption = Command29.Caption & iX & "." & HeadSellOnFixCol(iX) & " " '& ">" & GridX(ShowFixCol(iX)).GridIndexHead - 1 & ", " '[" & iX & ". " & GridX(iX).GridIndexHead & "] "
Next iX
papa:
'Command22.Caption = "000000"dad
End Sub

Sub HeadOnHSC()
'    ShowOnConHeadSell RangeEndHead
    If ConHeadSellOnFixCol > 0 Then
        Hsc.left = RangeEndHead + 7
        Hsc.Width = Picture1.ScaleWidth - RangeEndHead - 4
        
        HscHead.left = Picture1.left
        HscHead.Width = RangeEndHead + 3
        HscHead.top = Hsc.top
        
        HscHead.Visible = True
    Else
        Hsc.left = Picture1.left
        Hsc.Width = Picture1.Width
        
        HscHead.Visible = False
    End If
End Sub

Function LinePosisionHeads(Optional TmpGridRightXS As Integer) As Long
    If SellCountColumn - 1 = GridRightX And ShowOnConHeadSell - 1 = GridRightX - GridLeftX Then
        TmpGridRightXS = HeadSellOnFixCol(ShowOnConHeadSell - 1)
        LinePosisionHeads = vbRed
        If HeadCountAuto = False Then LineBack = GridX(ShowFixCol(TmpGridRightXS)).GridLeft + GridX(ShowFixCol(TmpGridRightXS)).GridWidth - RangeEndHead_Pos
    Else
        TmpGridRightXS = GridRightX
        If SellCountColumn - 1 = GridRightX Then LinePosisionHeads = vbYellow Else LinePosisionHeads = &HC0FFC0
    End If
End Function

Private Sub DrawingLineHeadCol(ObjMe As Object, nX1s As Integer, nRangeEndHeads As Integer, nColors As Long, Optional MyIndexY As Long = -1)
Dim LineY1 As Integer, LineY2 As Integer

If MyIndexY = -1 Then
    LineY1 = 0
    LineY2 = ObjMe.ScaleHeight
Else
    LineY1 = GridY(MyIndexY).GridTop
    LineY2 = GridY(MyIndexY).GridTop + GridY(MyIndexY).GridHeight
End If

ObjMe.Line (nX1s + 1, LineY1)- _
           (nX1s + nRangeEndHeads - 1, LineY2), Mix_Color(nColors, RGB(150, 150, 150)), BF
ObjMe.Line (nX1s + 3, LineY1)- _
           (nX1s + nRangeEndHeads - 1, LineY2), Mix_Color(nColors, RGB(50, 50, 50)), BF
End Sub

Sub Original_II_GridXY(ObjMe As Object, GridLeftX As Integer, GridUpY As Integer, Optional ShowPic As Boolean, Optional ShowX As Long = -1, Optional ShowY As Long = -1, Optional NoDrawing As Boolean)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim TFontWidth As Long, TMPConHeadSellOnFixCol As Integer
Dim ProsesShow As Integer, ProsesShowMin As Integer
Dim IndexHeadEnd As Integer, ShowConHeadSellOnFixCol As Integer
Dim TmpGridRightX As Integer
Dim TmpColorsLineHead As Long, CloseHeadFix As Boolean
Dim RangeEndHead As Integer, LineBack As Integer

TmpCutGridX = GridLeftX
TmpCutGridY = GridUpY

GridX(ShowFixCol(TmpCutGridX)).GridLeft = SetNewGrid.GridSize.GDRangeX
GridY(GridUpY).GridTop = SetNewGrid.GridSize.GDRangeY

ObjMe.Line (SetNewGrid.GridSize.GDRangeX, SetNewGrid.GridSize.GDRangeY)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), SetNewGrid.GridXYBackColor, BF

If SetNewGrid.GridStyleX = 0 Then
    ObjMe.Line (SetNewGrid.GridSize.GDRangeX, 0)- _
               (ObjMe.ScaleWidth, SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleX = 1 Or SetNewGrid.GridStyleX = 2 Then
    pvDrawGrid ObjMe, Val(SetNewGrid.GridSize.GDRangeX), 0, ObjMe.ScaleWidth, Val(SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), SetNewGrid.GridLenkapX.GridBackColGra(1), SetNewGrid.GridStyleX - 1
End If
If SetNewGrid.GridStyleY = 0 Then
    ObjMe.Line (0, SetNewGrid.GridSize.GDRangeY)- _
               (SetNewGrid.GridSize.GDRangeX, ObjMe.ScaleHeight), SetNewGrid.GridLenkapY.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleY = 1 Or SetNewGrid.GridStyleY = 2 Then
    pvDrawGrid ObjMe, 0, Val(SetNewGrid.GridSize.GDRangeY), Val(SetNewGrid.GridSize.GDRangeX), ObjMe.ScaleHeight, SetNewGrid.GridLenkapY.GridBackColGra(0), SetNewGrid.GridLenkapY.GridBackColGra(1), SetNewGrid.GridStyleY - 1
End If

''On Error GoTo papa
''ConHeadSellOnFixCol = -1
''If 8 = 8 Then
''    Static Vion As Integer
    
''    ConHeadSellOnFixCol = 3
''    ReDim Preserve HeadSellOnFixCol(ConHeadSellOnFixCol - 1)
''    HeadSellOnFixCol(0) = 0 '.IndexFixCol = 0
'''    GridX(HeadSellOnFixCol(0)).GridRealOnPosisi
'''    HeadSellOnFixCol(0).RealPosision = 0

''    HeadSellOnFixCol(1) = 2
''    HeadSellOnFixCol(2) = 10
'''    HeadSellOnFixCol(3) = 25
'''    HeadSellOnFixCol(4) = 28
'''    HeadSellOnFixCol(5) = 30
'''    HeadSellOnFixCol(6) = 35
''End If
'papax:
HeadCountAuto = False
RangeEndHead = 300 '>>>>>>>>>>>>>>>>

Command25.Caption = ""
TMPConHeadSellOnFixCol = 0 ' Vion
IndexHeadEnd = -1

ShowOnConHeadSell = 0
'GridX(*** ShowFixCol(TmpCutGridX)).GridLeft = 150
TMPConHeadSellOnFixCol = HscHead.Value
Do 'X
    ProsesShow = ShowFixCol(TmpCutGridX)
    If TmpCutGridX > GridLeftX Then ProsesShowMin = ShowFixCol(TmpCutGridX - 1)
    'If ShowConHeadSellOnFixCol > 2 Then CloseHeadFix = True
    If 8 = 8 Then
    If IndexHeadEnd <> -1 Then
        ProsesShowMin = IndexHeadEnd 'HeadSellOnFixCol(Vion)
        IndexHeadEnd = -1
    End If
    GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = 0

    If TMPConHeadSellOnFixCol < ConHeadSellOnFixCol Then
        If GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = -1 Then _
            GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        If GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = -1 Then _
            GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = ShowFixCol(TmpCutGridX)

        'If TMPConHeadSellOnFixCol > 2 + HscHead.Value Then CloseHeadFix = True
        If CloseHeadFix = False And GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi >= GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi Then
            If ShowFixCol(TmpCutGridX) <> HeadSellOnFixCol(TMPConHeadSellOnFixCol) Then
                ProsesShow = HeadSellOnFixCol(TMPConHeadSellOnFixCol)   '= 1
                GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridLeft = SetNewGrid.GridSize.GDRangeX  '+ 100
                
                If TmpCutGridX > GridLeftX Then ProsesShowMin = HeadSellOnFixCol(TMPConHeadSellOnFixCol - 1)
                ShowConHeadSellOnFixCol = ShowConHeadSellOnFixCol + 1
                ShowOnConHeadSell = ShowConHeadSellOnFixCol
                   
                GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = TMPConHeadSellOnFixCol + 1
                Command25.Caption = Command25.Caption & TmpCutGridX & " " & TMPConHeadSellOnFixCol & "-" & HeadSellOnFixCol(TMPConHeadSellOnFixCol) & ". "
            Else
            End If
            IndexHeadEnd = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        Else
            TMPConHeadSellOnFixCol = ConHeadSellOnFixCol
        End If
    Else
    End If
    
    TMPConHeadSellOnFixCol = TMPConHeadSellOnFixCol + 1
    End If
    '--------------
    If ShowConHeadSellOnFixCol > 0 And TmpCutGridX = GridLeftX + ShowConHeadSellOnFixCol Then
        RangeEndHead = 5
        'RangeEndHead = 250 '>>>>>>>>>>>>>>>>
    Else
        RangeEndHead = 0
        'RangeEndHead = 0
    End If
    If GridX(ProsesShow).GWidthDefault = False Then GridX(ProsesShow).GridWidth = SetNewGrid.GridSize.SellWidth_Def  '100
'    If HeadCountAuto = False Then
'MsgBox ProsesShow & " " & ProsesShowMin
    If TmpCutGridX > GridLeftX Then _
    GridX(ProsesShow).GridLeft = GridX(ProsesShowMin).GridLeft + GridX(ProsesShowMin).GridWidth + RangeEndHead '- RangeEndHead
    
'    If ShowConHeadSellOnFixCol = 3 And GridLeftX + ShowConHeadSellOnFixCol = ProsesShow Then
    If ShowConHeadSellOnFixCol > 0 And GridLeftX + ShowConHeadSellOnFixCol = ShowFixCol(ProsesShow) Then
        ObjMe.Line (RangeEndHead, 0)-(ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF
        GridX(ProsesShow).GridLeft = RangeEndHead + RangeEndHead
    End If

    If GridX(ProsesShow).GridLeft + GridX(ProsesShow).GridWidth >= RangeEndHead Then 'Menutup proses head
'        MsgBox ProsesShow
        CloseHeadFix = True
    End If
    
    Do 'Y
        If GridXYData(ProsesShow, TmpCutGridY).GridSub.GColorDefault(0) = False Then _
        GridXYData(ProsesShow, TmpCutGridY).GridSub.BackColor = SetNewGrid.GridXYBackColorSub
        If GridY(TmpCutGridY).GHeightDefault = False Then
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def
        Else
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def + GridY(TmpCutGridY).GHSave '20
        End If
        If TmpCutGridY > GridUpY Then _
        GridY(TmpCutGridY).GridTop = GridY(TmpCutGridY - 1).GridTop + GridY(TmpCutGridY - 1).GridHeight + 0
        
        If DoitY = False Then DrawDrid ObjMe, ProsesShow, TmpCutGridY, 1, , NoDrawing
        DrawDrid ObjMe, ProsesShow, TmpCutGridY, 0, , NoDrawing
        
        'UNTUK PENAMBAHAN CONTROL PADA SUB GRID
        If GridY(TmpCutGridY).GHeightDefault = True Then
            Select Case GridXYData(ProsesShow, TmpCutGridY).GridSubType.TypeControl '= "1" Then
            Case 1
                DrawList ProsesShow, TmpCutGridY, ObjMe ', Fg, Ad
'            Case 2
'                DrawButton *** ShowFixCol(TmpCutGridX), TmpCutGridY, ObjMe
            End Select
        End If
                
        If GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.TableHeight Or _
        TmpCutGridY + 1 = SetNewGrid.GridYCount Then
            GridDownY = TmpCutGridY '- 10
            TmpCutGridY = GridUpY

            Exit Do
        End If
    TmpCutGridY = TmpCutGridY + 1
    Loop

    DrawDrid ObjMe, ProsesShow, TmpCutGridY, 2, , NoDrawing
    
    If GridX(ProsesShow).GridLeft + GridX(ProsesShow).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    TmpCutGridX = SetNewGrid.GridXCount - HideCountFixCol - 1 Then
        GridRightX = TmpCutGridX  '*** ShowFixCol(TmpCutGridX) '- 10
        
        Exit Do
    End If
TmpCutGridX = TmpCutGridX + 1
DoitY = True
Loop

If ShowOnConHeadSell > 0 Then
    HeadWidFixCol = GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth
'    MsgBox GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth
Else
    HeadWidFixCol = 0
End If

ObjMe.Line (0, GridY(GridDownY).GridTop + GridY(GridDownY).GridHeight + 1)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

'>>>>>>>>>>>>> this are in eror &HC0FFC0
If SellCountColumn - 1 = GridRightX And ShowOnConHeadSell - 1 = GridRightX - GridLeftX Then
    TmpGridRightX = HeadSellOnFixCol(ShowOnConHeadSell - 1)
    TmpColorsLineHead = vbRed
    If HeadCountAuto = False Then LineBack = GridX(ShowFixCol(TmpGridRightX)).GridLeft + GridX(ShowFixCol(TmpGridRightX)).GridWidth - RangeEndHead
Else 'HeadCountAuto
    TmpGridRightX = GridRightX
    If SellCountColumn - 1 = GridRightX Then TmpColorsLineHead = vbYellow Else TmpColorsLineHead = &HC0FFC0
End If

ObjMe.Line (GridX(ShowFixCol(TmpGridRightX)).GridLeft + GridX(ShowFixCol(TmpGridRightX)).GridWidth + 1 - LineBack, 0)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

If ShowOnConHeadSell > 0 Then
    If HeadCountAuto = False Then
        DrawingLineHeadCol ObjMe, _
        RangeEndHead, 5, TmpColorsLineHead
    Else
        DrawingLineHeadCol ObjMe, _
        GridX(HeadSellOnFixCol(ShowConHeadSellOnFixCol - 1 + HscHead.Value)).GridLeft + GridX(HeadSellOnFixCol(ShowConHeadSellOnFixCol - 1 + HscHead.Value)).GridWidth, 5, TmpColorsLineHead
    End If
End If

If Text1.Visible = True Then _
AllText Text1, SellText(FixdColnIndex(XIndexPointerText1), YIndexPointerText1), SellLeft(FixdColnIndex(XIndexPointerText1)), SellTop(YIndexPointerText1), SellWidth(FixdColnIndex(XIndexPointerText1)), SellHeight_Def
''On AllText Text1, SellText(XIndexPointerText1, YIndexPointerText1), SellLeft(XIndexPointerText1), SellTop(YIndexPointerText1), SellWidth(XIndexPointerText1), SellHeight_Def

'HscHead.Value >>>>>>>

Command24.Caption = ShowOnConHeadSell & " " & GridLeftX & " " & GridRightX - GridLeftX

'Command23.Caption = " "
'For iX = 0 To 9
'    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridLeft & ", "
'Next iX

'Exit Sub
Command24.Caption = ShowConHeadSellOnFixCol & "."
On Error GoTo papa
Command22.Caption = "HeadSellOnFixCol "
For iX = 0 To ConHeadSellOnFixCol - 1
    Command22.Caption = Command22.Caption & GridX(HeadSellOnFixCol(iX)).GridIndexHead - 1 & "." & HeadSellOnFixCol(iX) & " "
Next iX

Command23.Caption = "GridRealOnPosisi "
For iX = 0 To 12
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealOnPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridRealPosisi "
For iX = 0 To 12
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "ShowFixCol [" & UBound(ShowFixCol()) & "] "
For iX = 0 To 12 'UBound(ShowFixCol())
    Command23.Caption = Command23.Caption & iX & "." & ShowFixCol(iX) & " "  '& "." & GridX(ShowFixCol(iX)).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridIndexHead "
For iX = 0 To 12 'UBound(ShowFixCol())
    If GridX(ShowFixCol(iX)).GridIndexHead <> 0 Then
    Command23.Caption = Command23.Caption & iX & "." & ShowFixCol(iX) & ">" & GridX(ShowFixCol(iX)).GridIndexHead - 1 & ", " '[" & iX & ". " & GridX(iX).GridIndexHead & "] "
    End If
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridOnIndexHead "
For iX = 0 To 12 'UBound(ShowFixCol())
    'If GridX(ShowFixCol(iX)).GridIndexHead <> 0 Then
    Command23.Caption = Command23.Caption & iX & "." & GridX(ShowFixCol(iX)).GridOnIndexHead & ", "
    'End If UBound(ShowFixCol())
Next iX
papa:
'Command22.Caption = "000000"dad
End Sub

Sub Original_True_GridXY(ObjMe As Object, GridLeftX As Integer, GridUpY As Integer, Optional ShowPic As Boolean, Optional ShowX As Long = -1, Optional ShowY As Long = -1, Optional NoDrawing As Boolean)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim TFontWidth As Long, TMPConHeadSellOnFixCol As Integer
Dim ProsesShow As Integer, ProsesShowMin As Integer
Dim IndexHeadEnd As Integer, ShowConHeadSellOnFixCol As Integer
Dim RangeEndHead As Integer, TmpGridRightX As Integer
Dim TmpColorsLineHead As Long, CloseHeadFix As Boolean

TmpCutGridX = GridLeftX
TmpCutGridY = GridUpY

GridX(ShowFixCol(TmpCutGridX)).GridLeft = SetNewGrid.GridSize.GDRangeX
GridY(GridUpY).GridTop = SetNewGrid.GridSize.GDRangeY

ObjMe.Line (SetNewGrid.GridSize.GDRangeX, SetNewGrid.GridSize.GDRangeY)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), SetNewGrid.GridXYBackColor, BF

If SetNewGrid.GridStyleX = 0 Then
    ObjMe.Line (SetNewGrid.GridSize.GDRangeX, 0)- _
               (ObjMe.ScaleWidth, SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleX = 1 Or SetNewGrid.GridStyleX = 2 Then
    pvDrawGrid ObjMe, Val(SetNewGrid.GridSize.GDRangeX), 0, ObjMe.ScaleWidth, Val(SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), SetNewGrid.GridLenkapX.GridBackColGra(1), SetNewGrid.GridStyleX - 1
End If
If SetNewGrid.GridStyleY = 0 Then
    ObjMe.Line (0, SetNewGrid.GridSize.GDRangeY)- _
               (SetNewGrid.GridSize.GDRangeX, ObjMe.ScaleHeight), SetNewGrid.GridLenkapY.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleY = 1 Or SetNewGrid.GridStyleY = 2 Then
    pvDrawGrid ObjMe, 0, Val(SetNewGrid.GridSize.GDRangeY), Val(SetNewGrid.GridSize.GDRangeX), ObjMe.ScaleHeight, SetNewGrid.GridLenkapY.GridBackColGra(0), SetNewGrid.GridLenkapY.GridBackColGra(1), SetNewGrid.GridStyleY - 1
End If

''On Error GoTo papa
''ConHeadSellOnFixCol = -1
''If 8 = 8 Then
''    Static Vion As Integer
    
''    ConHeadSellOnFixCol = 3
''    ReDim Preserve HeadSellOnFixCol(ConHeadSellOnFixCol - 1)
''    HeadSellOnFixCol(0) = 0 '.IndexFixCol = 0
'''    GridX(HeadSellOnFixCol(0)).GridRealOnPosisi
'''    HeadSellOnFixCol(0).RealPosision = 0

''    HeadSellOnFixCol(1) = 2
''    HeadSellOnFixCol(2) = 10
'''    HeadSellOnFixCol(3) = 25
'''    HeadSellOnFixCol(4) = 28
'''    HeadSellOnFixCol(5) = 30
'''    HeadSellOnFixCol(6) = 35
''End If
'papax:

Command25.Caption = ""
TMPConHeadSellOnFixCol = 0 ' Vion
IndexHeadEnd = -1

ShowOnConHeadSell = 0
'GridX(*** ShowFixCol(TmpCutGridX)).GridLeft = 150
TMPConHeadSellOnFixCol = HscHead.Value
Do 'X
    ProsesShow = ShowFixCol(TmpCutGridX)
    If TmpCutGridX > GridLeftX Then ProsesShowMin = ShowFixCol(TmpCutGridX - 1)
    'If ShowConHeadSellOnFixCol > 2 Then CloseHeadFix = True
    If 8 = 8 Then
    If IndexHeadEnd <> -1 Then
        ProsesShowMin = IndexHeadEnd 'HeadSellOnFixCol(Vion)
        IndexHeadEnd = -1
    End If
    GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = 0

    If TMPConHeadSellOnFixCol < ConHeadSellOnFixCol Then
    
        If GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = -1 Then _
            GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        If GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = -1 Then _
            GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = ShowFixCol(TmpCutGridX)
                                    ...........
        If TMPConHeadSellOnFixCol > 2 + HscHead.Value Then CloseHeadFix = True
        If CloseHeadFix = False And GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi >= GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi Then
            If ShowFixCol(TmpCutGridX) <> HeadSellOnFixCol(TMPConHeadSellOnFixCol) Then
                ProsesShow = HeadSellOnFixCol(TMPConHeadSellOnFixCol)   '= 1
                GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridLeft = SetNewGrid.GridSize.GDRangeX  '+ 100
                
                If TmpCutGridX > GridLeftX Then ProsesShowMin = HeadSellOnFixCol(TMPConHeadSellOnFixCol - 1)
                ShowConHeadSellOnFixCol = ShowConHeadSellOnFixCol + 1
                ShowOnConHeadSell = ShowConHeadSellOnFixCol
                   
                GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = TMPConHeadSellOnFixCol + 1
                Command25.Caption = Command25.Caption & TmpCutGridX & " " & TMPConHeadSellOnFixCol & "-" & HeadSellOnFixCol(TMPConHeadSellOnFixCol) & ". "
            Else
            End If
            IndexHeadEnd = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        Else
            TMPConHeadSellOnFixCol = ConHeadSellOnFixCol
        End If
    Else
    End If
    TMPConHeadSellOnFixCol = TMPConHeadSellOnFixCol + 1
    End If
    '--------------
    If ShowConHeadSellOnFixCol > 0 And TmpCutGridX = GridLeftX + ShowConHeadSellOnFixCol Then
        RangeEndHead = 5
    Else
        RangeEndHead = 0
    End If
    If GridX(ProsesShow).GWidthDefault = False Then GridX(ProsesShow).GridWidth = SetNewGrid.GridSize.SellWidth_Def  '100
    If TmpCutGridX > GridLeftX Then _
    GridX(ProsesShow).GridLeft = GridX(ProsesShowMin).GridLeft + GridX(ProsesShowMin).GridWidth + RangeEndHead

    Do 'Y
        If GridXYData(ProsesShow, TmpCutGridY).GridSub.GColorDefault(0) = False Then _
        GridXYData(ProsesShow, TmpCutGridY).GridSub.BackColor = SetNewGrid.GridXYBackColorSub
        If GridY(TmpCutGridY).GHeightDefault = False Then
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def
        Else
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def + GridY(TmpCutGridY).GHSave '20
        End If
        If TmpCutGridY > GridUpY Then _
        GridY(TmpCutGridY).GridTop = GridY(TmpCutGridY - 1).GridTop + GridY(TmpCutGridY - 1).GridHeight + 0
        
        If DoitY = False Then DrawDrid ObjMe, ProsesShow, TmpCutGridY, 1, , NoDrawing
        DrawDrid ObjMe, ProsesShow, TmpCutGridY, 0, , NoDrawing
        
        'UNTUK PENAMBAHAN CONTROL PADA SUB GRID
        If GridY(TmpCutGridY).GHeightDefault = True Then
            Select Case GridXYData(ProsesShow, TmpCutGridY).GridSubType.TypeControl '= "1" Then
            Case 1
                DrawList ProsesShow, TmpCutGridY, ObjMe ', Fg, Ad
'            Case 2
'                DrawButton *** ShowFixCol(TmpCutGridX), TmpCutGridY, ObjMe
            End Select
        End If
                
        If GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.TableHeight Or _
        TmpCutGridY + 1 = SetNewGrid.GridYCount Then
            GridDownY = TmpCutGridY '- 10
            TmpCutGridY = GridUpY

            Exit Do
        End If
    TmpCutGridY = TmpCutGridY + 1
    Loop

    DrawDrid ObjMe, ProsesShow, TmpCutGridY, 2, , NoDrawing
    
    If GridX(ProsesShow).GridLeft + GridX(ProsesShow).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    TmpCutGridX = SetNewGrid.GridXCount - HideCountFixCol - 1 Then
        GridRightX = TmpCutGridX  '*** ShowFixCol(TmpCutGridX) '- 10
        
        Exit Do
    End If
TmpCutGridX = TmpCutGridX + 1
DoitY = True
Loop

If ShowOnConHeadSell > 0 Then
    HeadWidFixCol = GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth
'    MsgBox GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth
Else
    HeadWidFixCol = 0
End If

ObjMe.Line (0, GridY(GridDownY).GridTop + GridY(GridDownY).GridHeight + 1)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

'>>>>>>>>>>>>> this are in eror &HC0FFC0
If SellCountColumn - 1 = GridRightX And ShowOnConHeadSell - 1 = GridRightX - GridLeftX Then
    TmpGridRightX = HeadSellOnFixCol(ShowOnConHeadSell - 1)
    TmpColorsLineHead = vbRed
Else
    TmpGridRightX = GridRightX
    If SellCountColumn - 1 = GridRightX Then TmpColorsLineHead = vbYellow Else TmpColorsLineHead = &HC0FFC0
End If
ObjMe.Line (GridX(ShowFixCol(TmpGridRightX)).GridLeft + GridX(ShowFixCol(TmpGridRightX)).GridWidth + 1, 0)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF
If ShowOnConHeadSell > 0 Then DrawingLineHeadCol ObjMe, GridX(HeadSellOnFixCol(ShowConHeadSellOnFixCol - 1 + HscHead.Value)).GridLeft + GridX(HeadSellOnFixCol(ShowConHeadSellOnFixCol - 1 + HscHead.Value)).GridWidth, 5, TmpColorsLineHead

If Text1.Visible = True Then _
AllText Text1, SellText(FixdColnIndex(XIndexPointerText1), YIndexPointerText1), SellLeft(FixdColnIndex(XIndexPointerText1)), SellTop(YIndexPointerText1), SellWidth(FixdColnIndex(XIndexPointerText1)), SellHeight_Def
''On AllText Text1, SellText(XIndexPointerText1, YIndexPointerText1), SellLeft(XIndexPointerText1), SellTop(YIndexPointerText1), SellWidth(XIndexPointerText1), SellHeight_Def

'HscHead.Value >>>>>>>

Command24.Caption = ShowOnConHeadSell & " " & GridLeftX & " " & GridRightX - GridLeftX

'Command23.Caption = " "
'For iX = 0 To 9
'    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridLeft & ", "
'Next iX

Exit Sub
Command24.Caption = ShowConHeadSellOnFixCol & "."
On Error GoTo papa
Command22.Caption = "HeadSellOnFixCol "
For iX = 0 To ConHeadSellOnFixCol - 1
    Command22.Caption = Command22.Caption & GridX(HeadSellOnFixCol(iX)).GridIndexHead - 1 & "." & HeadSellOnFixCol(iX) & " "
Next iX

Command23.Caption = "GridRealOnPosisi "
For iX = 0 To 9
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealOnPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridRealPosisi "
For iX = 0 To 12
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "ShowFixCol [" & UBound(ShowFixCol()) & "] "
For iX = 0 To 12 'UBound(ShowFixCol())
    Command23.Caption = Command23.Caption & iX & "." & ShowFixCol(iX) & " "  '& "." & GridX(ShowFixCol(iX)).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridIndexHead "
For iX = 0 To 12 'UBound(ShowFixCol())
    If GridX(ShowFixCol(iX)).GridIndexHead <> 0 Then
    Command23.Caption = Command23.Caption & iX & "." & ShowFixCol(iX) & ">" & GridX(ShowFixCol(iX)).GridIndexHead - 1 & ", " '[" & iX & ". " & GridX(iX).GridIndexHead & "] "
    End If
Next iX
Command23.Caption = Command23.Caption & vbCrLf & "GridOnIndexHead "
For iX = 0 To 12 'UBound(ShowFixCol())
    'If GridX(ShowFixCol(iX)).GridIndexHead <> 0 Then
    Command23.Caption = Command23.Caption & iX & "." & GridX(ShowFixCol(iX)).GridOnIndexHead & ", "
    'End If UBound(ShowFixCol())
Next iX
papa:
'Command22.Caption = "000000"
End Sub

Sub TMPGridXY(ObjMe As Object, GridLeftX As Integer, GridUpY As Integer, Optional ShowPic As Boolean, Optional ShowX As Long = -1, Optional ShowY As Long = -1, Optional NoDrawing As Boolean)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim TFontWidth As Long, TMPConHeadSellOnFixCol As Integer
Dim ProsesShow As Integer, ProsesShowMin As Integer
Dim IndexHeadEnd As Integer, ShowConHeadSellOnFixCol As Integer
Dim RangeEndHead As Integer

TmpCutGridX = GridLeftX
TmpCutGridY = GridUpY

GridX(ShowFixCol(TmpCutGridX)).GridLeft = SetNewGrid.GridSize.GDRangeX
GridY(GridUpY).GridTop = SetNewGrid.GridSize.GDRangeY

ObjMe.Line (SetNewGrid.GridSize.GDRangeX, SetNewGrid.GridSize.GDRangeY)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), SetNewGrid.GridXYBackColor, BF

If SetNewGrid.GridStyleX = 0 Then
    ObjMe.Line (SetNewGrid.GridSize.GDRangeX, 0)- _
               (ObjMe.ScaleWidth, SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleX = 1 Or SetNewGrid.GridStyleX = 2 Then
    pvDrawGrid ObjMe, Val(SetNewGrid.GridSize.GDRangeX), 0, ObjMe.ScaleWidth, Val(SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), SetNewGrid.GridLenkapX.GridBackColGra(1), SetNewGrid.GridStyleX - 1
End If
If SetNewGrid.GridStyleY = 0 Then
    ObjMe.Line (0, SetNewGrid.GridSize.GDRangeY)- _
               (SetNewGrid.GridSize.GDRangeX, ObjMe.ScaleHeight), SetNewGrid.GridLenkapY.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleY = 1 Or SetNewGrid.GridStyleY = 2 Then
    pvDrawGrid ObjMe, 0, Val(SetNewGrid.GridSize.GDRangeY), Val(SetNewGrid.GridSize.GDRangeX), ObjMe.ScaleHeight, SetNewGrid.GridLenkapY.GridBackColGra(0), SetNewGrid.GridLenkapY.GridBackColGra(1), SetNewGrid.GridStyleY - 1
End If

Do 'X
    
    If 8 = 8 Then
    If GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead > 0 Then
        ShowFixCol(TmpCutGridX) = GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead - 1
        GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = 0
    End If
    
    If TMPConHeadSellOnFixCol < ConHeadSellOnFixCol Then
        If GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = -1 Then _
            GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        If GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = -1 Then _
            GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi = ShowFixCol(TmpCutGridX)
        
        If GridX(ShowFixCol(TmpCutGridX)).GridRealOnPosisi >= GridX(HeadSellOnFixCol(TMPConHeadSellOnFixCol)).GridRealOnPosisi Then
            GridX(ShowFixCol(TmpCutGridX)).GridOnIndexHead = ShowFixCol(TmpCutGridX) + 1
            ShowFixCol(TmpCutGridX) = HeadSellOnFixCol(TMPConHeadSellOnFixCol)
        End If
    End If
    TMPConHeadSellOnFixCol = TMPConHeadSellOnFixCol + 1
    End If
    
    If GridX(ShowFixCol(TmpCutGridX)).GWidthDefault = False Then GridX(ShowFixCol(TmpCutGridX)).GridWidth = SetNewGrid.GridSize.SellWidth_Def  '100
    If TmpCutGridX > GridLeftX Then _
    GridX(ShowFixCol(TmpCutGridX)).GridLeft = GridX(ShowFixCol(TmpCutGridX - 1) - 0).GridLeft + GridX(ShowFixCol(TmpCutGridX - 1) - 0).GridWidth

    Do 'Y
'        GridXYData(ShowFixCol(TmpCutGridX), TmpCutGridY).GridXYValue = "Texts"
'
        If GridXYData(ShowFixCol(TmpCutGridX), TmpCutGridY).GridSub.GColorDefault(0) = False Then _
        GridXYData(ShowFixCol(TmpCutGridX), TmpCutGridY).GridSub.BackColor = SetNewGrid.GridXYBackColorSub
        If GridY(TmpCutGridY).GHeightDefault = False Then
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def
        Else
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def + GridY(TmpCutGridY).GHSave '20
        End If
        If TmpCutGridY > GridUpY Then _
        GridY(TmpCutGridY).GridTop = GridY(TmpCutGridY - 1).GridTop + GridY(TmpCutGridY - 1).GridHeight
        
        
        'If GridX(ShowFixCol(TmpCutGridX)).Visibles = True Then GridX(ShowFixCol(TmpCutGridX)).GridWidth = 0
        
        If DoitY = False Then DrawDrid ObjMe, ShowFixCol(TmpCutGridX), TmpCutGridY, 1, , NoDrawing
        DrawDrid ObjMe, ShowFixCol(TmpCutGridX), TmpCutGridY, 0, , NoDrawing
        
        'UNTUK PENAMBAHAN CONTROL PADA SUB GRID
        If GridY(TmpCutGridY).GHeightDefault = True Then
            Select Case GridXYData(ShowFixCol(TmpCutGridX), TmpCutGridY).GridSubType.TypeControl '= "1" Then
            Case 1
                DrawList ShowFixCol(TmpCutGridX), TmpCutGridY, ObjMe ', Fg, Ad
'            Case 2
'                DrawButton ShowFixCol(TmpCutGridX), TmpCutGridY, ObjMe
            End Select
        End If
        '---------------------
                
        'End If
                
        If GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.TableHeight Or _
        TmpCutGridY + 1 = SetNewGrid.GridYCount Then
            GridDownY = TmpCutGridY '- 10
            TmpCutGridY = GridUpY

            Exit Do
        End If
    TmpCutGridY = TmpCutGridY + 1
    Loop

    DrawDrid ObjMe, ShowFixCol(TmpCutGridX), TmpCutGridY, 2, , NoDrawing
    
    
'    If GridX(ShowFixCol(TmpCutGridX)).GridLeft + GridX(ShowFixCol(TmpCutGridX)).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    ShowFixCol(TmpCutGridX) + 1 = SetNewGrid.GridXCount Then
'        GridRightX = TmpCutGridX 'ShowFixCol(TmpCutGridX) '- 10
    
'    If GridX(ShowFixCol(TmpCutGridX)).GridLeft + GridX(ShowFixCol(TmpCutGridX)).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    ShowFixCol(TmpCutGridX) + 1 = SetNewGrid.GridXCount - HideCountFixCol Then
    If GridX(ShowFixCol(TmpCutGridX)).GridLeft + GridX(ShowFixCol(TmpCutGridX)).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    TmpCutGridX = SetNewGrid.GridXCount - HideCountFixCol - 1 Then
        GridRightX = TmpCutGridX  'ShowFixCol(TmpCutGridX) '- 10
        
        Exit Do
    End If
'If GridX(ShowFixCol(TmpCutGridX)).Visibles = True Then ShowFixCol(TmpCutGridX) = ShowFixCol(TmpCutGridX) + 1
TmpCutGridX = TmpCutGridX + 1
DoitY = True
        
''    If ShowFixCol(TmpCutGridX) = 1 And GridLeftX > 0 Then ------------------------------------
''        ShowFixCol(TmpCutGridX) = GridLeftX                                                   |
''        GridX(ShowFixCol(TmpCutGridX)).GridLeft = GridX(0).GridWidth + GridX(0).GridLeft + 3  |-- This Tester
''        ShowX = -1                                                                |
''    End If -----------------------------------------------------------------------
Loop

ObjMe.Line (0, GridY(GridDownY).GridTop + GridY(GridDownY).GridHeight + 1)- _
(ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

ObjMe.Line (GridX(ShowFixCol(GridRightX)).GridLeft + GridX(ShowFixCol(GridRightX)).GridWidth + 1, 0)- _
(ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

If Text1.Visible = True Then _
AllText Text1, SellText(XIndexPointerText1, YIndexPointerText1), SellLeft(XIndexPointerText1), SellTop(YIndexPointerText1), SellWidth(XIndexPointerText1), SellHeight_Def

Command24.Caption = ShowConHeadSellOnFixCol
On Error GoTo papa
Command22.Caption = ""
For iX = 0 To ConHeadSellOnFixCol - 1
    Command22.Caption = Command22.Caption & GridX(HeadSellOnFixCol(iX)).GridIndexHead - 1 & "." & HeadSellOnFixCol(iX) & " "
Next iX

Command23.Caption = ""
For iX = 0 To 9
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealOnPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf
For iX = 0 To 9
    Command23.Caption = Command23.Caption & iX & "." & GridX(iX).GridRealPosisi & ", "
Next iX
Command23.Caption = Command23.Caption & vbCrLf
For iX = 0 To UBound(ShowFixCol())
    Command23.Caption = Command23.Caption & ShowFixCol(iX) & "." & GridX(ShowFixCol(iX)).GridRealPosisi & ", "
Next iX
papa:
End Sub

Sub UPDATE1_GridXY(ObjMe As Object, GridLeftX As Integer, GridUpY As Integer, Optional ShowPic As Boolean, Optional ShowX As Long = -1, Optional ShowY As Long = -1, Optional NoDrawing As Boolean)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim TFontWidth As Long

''Dim ShowFixCol(100) As Integer
''Dim iY As Integer, iYs As Integer

''For iY = 0 To 100
''    iYs = iY
'    If iY = 1 Then iYs = iYs + 1
''ShowFixCol(iY) = iYs
''Next iY
''ShowFixCol(0) = 0
''ShowFixCol(1) = 3
''ShowFixCol(2) = 2
''ShowFixCol(3) = 4
''ShowFixCol(4) = 5
''ShowFixCol(5) = 6
''ShowFixCol(6) = 7
''ShowFixCol(7) = 8
''ShowFixCol(8) = 9
''ShowFixCol(9) = 10
''ShowFixCol(10) = 11
''ShowFixCol(11) = 12
''ShowFixCol(12) = 13

'ObjMe.Picture = Nothing
'ObjMe.Cls
'ObjMe.BackColor = 0

TmpCutGridX = GridLeftX
TmpCutGridY = GridUpY

GridX(ShowFixCol(TmpCutGridX)).GridLeft = SetNewGrid.GridSize.GDRangeX
GridY(GridUpY).GridTop = SetNewGrid.GridSize.GDRangeY

'Reset SetNewGrid.GridSize.SellHeight_Def
'Form1.AutoRedraw

'Form1.AutoRedraw
'ObjMe.AutoRedraw = False
    'Picture1.Picture = Nothing
    'Picture1.BackColor = SetNewGrid.GridXYBackColor
ObjMe.Line (SetNewGrid.GridSize.GDRangeX, SetNewGrid.GridSize.GDRangeY)- _
           (ObjMe.ScaleWidth, ObjMe.ScaleHeight), SetNewGrid.GridXYBackColor, BF

If SetNewGrid.GridStyleX = 0 Then
    ObjMe.Line (SetNewGrid.GridSize.GDRangeX, 0)- _
               (ObjMe.ScaleWidth, SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleX = 1 Or SetNewGrid.GridStyleX = 2 Then
    pvDrawGrid ObjMe, Val(SetNewGrid.GridSize.GDRangeX), 0, ObjMe.ScaleWidth, Val(SetNewGrid.GridSize.GDRangeY), SetNewGrid.GridLenkapX.GridBackColGra(0), SetNewGrid.GridLenkapX.GridBackColGra(1), SetNewGrid.GridStyleX - 1
End If
If SetNewGrid.GridStyleY = 0 Then
    ObjMe.Line (0, SetNewGrid.GridSize.GDRangeY)- _
               (SetNewGrid.GridSize.GDRangeX, ObjMe.ScaleHeight), SetNewGrid.GridLenkapY.GridBackColGra(0), BF
ElseIf SetNewGrid.GridStyleY = 1 Or SetNewGrid.GridStyleY = 2 Then
    pvDrawGrid ObjMe, 0, Val(SetNewGrid.GridSize.GDRangeY), Val(SetNewGrid.GridSize.GDRangeX), ObjMe.ScaleHeight, SetNewGrid.GridLenkapY.GridBackColGra(0), SetNewGrid.GridLenkapY.GridBackColGra(1), SetNewGrid.GridStyleY - 1
End If


'ShowX = 2
'ShowY = 150
'ShowFixCol(TmpCutGridX) = 2
'MsgBox GridX(1).GridLeft
'GridX(ShowFixCol(TmpCutGridX)).GridLeft = GridX(0).GridWidth
'ShowFixCol(TmpCutGridX) = ShowFixCol(TmpCutGridX)
'If ShowX > -1 And ShowX = ShowFixCol(TmpCutGridX) Then Exit Do

'If ShowX > -1 And GridLeftX <> 0 Then
''    ShowFixCol(TmpCutGridX) = 0 This Tester
'End If

'GridX(0).GridFront = True
'GridX(ShowFixCol(TmpCutGridX)).GridLeft = GridX(0).GridWidth + 25 ' + 100
Dim FrontSell(1) As Integer
If 8 = 9 Then
FrontSell(0) = 5
'FixdColnVisible(3) = True
FrontSell(1) = 2
'FixdColnVisible(2) = True
For iX = 0 To 1
    'ShowFixCol(TmpCutGridX + iX) = FrontSell(iX)
Next iX
End If
Command22.Caption = TmpCutGridX

Do 'X
'    If GridX(ShowFixCol(TmpCutGridX)).GridFront = True Then
                    
'    End If
'        Command22.Caption = TmpCutGridX
    'MsgBox ShowFixCol(TmpCutGridX)
    'ShowFixCol(TmpCutGridX) = 3
    If GridX(ShowFixCol(TmpCutGridX)).GWidthDefault = False Then GridX(ShowFixCol(TmpCutGridX)).GridWidth = SetNewGrid.GridSize.SellWidth_Def  '100
    If TmpCutGridX > GridLeftX Then _
    GridX(ShowFixCol(TmpCutGridX)).GridLeft = GridX(ShowFixCol(TmpCutGridX - 1) - 0).GridLeft + GridX(ShowFixCol(TmpCutGridX - 1) - 0).GridWidth

    Do 'Y
'        GridXYData(ShowFixCol(TmpCutGridX), TmpCutGridY).GridXYValue = "Texts"
'
        If GridXYData(ShowFixCol(TmpCutGridX), TmpCutGridY).GridSub.GColorDefault(0) = False Then _
        GridXYData(ShowFixCol(TmpCutGridX), TmpCutGridY).GridSub.BackColor = SetNewGrid.GridXYBackColorSub
        If GridY(TmpCutGridY).GHeightDefault = False Then
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def
        Else
            GridY(TmpCutGridY).GridHeight = SetNewGrid.GridSize.SellHeight_Def + GridY(TmpCutGridY).GHSave '20
        End If
        If TmpCutGridY > GridUpY Then _
        GridY(TmpCutGridY).GridTop = GridY(TmpCutGridY - 1).GridTop + GridY(TmpCutGridY - 1).GridHeight
        
        
        'If GridX(ShowFixCol(TmpCutGridX)).Visibles = True Then GridX(ShowFixCol(TmpCutGridX)).GridWidth = 0
        
        If DoitY = False Then DrawDrid ObjMe, ShowFixCol(TmpCutGridX), TmpCutGridY, 1, , NoDrawing
        DrawDrid ObjMe, ShowFixCol(TmpCutGridX), TmpCutGridY, 0, , NoDrawing
        
        'UNTUK PENAMBAHAN CONTROL PADA SUB GRID
        If GridY(TmpCutGridY).GHeightDefault = True Then
            Select Case GridXYData(ShowFixCol(TmpCutGridX), TmpCutGridY).GridSubType.TypeControl '= "1" Then
            Case 1
                DrawList ShowFixCol(TmpCutGridX), TmpCutGridY, ObjMe ', Fg, Ad
'            Case 2
'                DrawButton ShowFixCol(TmpCutGridX), TmpCutGridY, ObjMe
            End Select
        End If
        '---------------------
                
        'End If
                
        If GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.TableHeight Or _
        TmpCutGridY + 1 = SetNewGrid.GridYCount Then
            GridDownY = TmpCutGridY '- 10
            TmpCutGridY = GridUpY

            Exit Do
        End If
    TmpCutGridY = TmpCutGridY + 1
    Loop

    DrawDrid ObjMe, ShowFixCol(TmpCutGridX), TmpCutGridY, 2, , NoDrawing
    
    
'    If GridX(ShowFixCol(TmpCutGridX)).GridLeft + GridX(ShowFixCol(TmpCutGridX)).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    ShowFixCol(TmpCutGridX) + 1 = SetNewGrid.GridXCount Then
'        GridRightX = TmpCutGridX 'ShowFixCol(TmpCutGridX) '- 10
    
'    If GridX(ShowFixCol(TmpCutGridX)).GridLeft + GridX(ShowFixCol(TmpCutGridX)).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    ShowFixCol(TmpCutGridX) + 1 = SetNewGrid.GridXCount - HideCountFixCol Then
    If GridX(ShowFixCol(TmpCutGridX)).GridLeft + GridX(ShowFixCol(TmpCutGridX)).GridWidth >= SetNewGrid.GridSize.TableWidth Or _
    TmpCutGridX = SetNewGrid.GridXCount - HideCountFixCol - 1 Then
        GridRightX = TmpCutGridX  'ShowFixCol(TmpCutGridX) '- 10
        
        Exit Do
    End If
'If GridX(ShowFixCol(TmpCutGridX)).Visibles = True Then ShowFixCol(TmpCutGridX) = ShowFixCol(TmpCutGridX) + 1
TmpCutGridX = TmpCutGridX + 1
DoitY = True
        
''    If ShowFixCol(TmpCutGridX) = 1 And GridLeftX > 0 Then ------------------------------------
''        ShowFixCol(TmpCutGridX) = GridLeftX                                                   |
''        GridX(ShowFixCol(TmpCutGridX)).GridLeft = GridX(0).GridWidth + GridX(0).GridLeft + 3  |-- This Tester
''        ShowX = -1                                                                |
''    End If -----------------------------------------------------------------------
Loop

ObjMe.Line (0, GridY(GridDownY).GridTop + GridY(GridDownY).GridHeight + 1)- _
(ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

ObjMe.Line (GridX(ShowFixCol(GridRightX)).GridLeft + GridX(ShowFixCol(GridRightX)).GridWidth + 1, 0)- _
(ObjMe.ScaleWidth, ObjMe.ScaleHeight), ObjMe.BackColor, BF

If Text1.Visible = True Then _
AllText Text1, SellText(XIndexPointerText1, YIndexPointerText1), SellLeft(XIndexPointerText1), SellTop(YIndexPointerText1), SellWidth(XIndexPointerText1), SellHeight_Def

'Text1.left = GridX(XIndexPointerText1).GridLeft
'Text1.top = GridY(YIndexPointerText1).GridTop
'Form1.ScaleWidth
'ObjMe.Picture = ObjMe.Image
'ObjMe.AutoRedraw = True

'MsgBox TmpCutGridX - Hsc.Value & " " & GridRightX
End Sub

Sub DrawDrid(ObjMe As Object, TmpCutGridX As Integer, TmpCutGridY As Integer, TypeDraw As Integer, Optional Pilih As Integer, Optional NoDrawing As Boolean, Optional OnlyLineHead As Boolean)
Dim TmpRangePicGridX1 As Integer, TmpRangePicGridX2 As Integer
Dim tmpColorGrid As Long, tmpColorGrid2 As Long
Dim Oles(1) As OLE_COLOR, OverRangeEndHeads As Integer
Dim TmpColorsLineHead As Long

'** tambahkan pada
'   1. option drawing untuk mempersingkat bahwa line ini di gunakan bukan pada GridXY
'   2. sempurnakan pada gambar line di RangeEndHead
If NoDrawing = True Then Exit Sub
If GridX(TmpCutGridX).GridLeft < 0 Then Exit Sub


Cx = GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 0
Cy = GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight

DefText ObjMe
Select Case TypeDraw
Case 0
    If HeadCountAuto = False And ShowOnConHeadSell > 0 And GridX(TmpCutGridX).GridIndexHead = ShowOnConHeadSell + HscHead.Value And GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth > RangeEndHead Then
        OverRangeEndHeads = GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 1 - RangeEndHead
    Else
        OverRangeEndHeads = 0
    End If
    
    If Pilih = 0 Or Pilih = 1 Then

        If GridXYData(TmpCutGridX, TmpCutGridY).Grid.GColorDefault(0) = True Then _
           ObjMe.Line (GridX(TmpCutGridX).GridLeft + 1, GridY(TmpCutGridY).GridTop + 0)- _
           (Cx - 1 - OverRangeEndHeads, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.SellHeight_Def), GridXYData(TmpCutGridX, TmpCutGridY).Grid.BackColor, BF '  &H8000000F, BF
        
        If 8 = 8 Then
        ObjMe.Line (GridX(TmpCutGridX).GridLeft - 0, GridY(TmpCutGridY).GridTop + 2)- _
                      (Cx - 2 - OverRangeEndHeads, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.SellHeight_Def), vbWhite, B
        End If
        
        If GridX(TmpCutGridX).GPicturePut = True Then
            DrawPicGrid ObjMe, GridXYData(TmpCutGridX, TmpCutGridY).GridXYPicIndex, _
            TmpCutGridX, TmpCutGridY
    
            TmpRangePicGridX1 = SetNewGrid.GridSizePic.SellIconX1
            TmpRangePicGridX2 = SetNewGrid.GridSizePic.SellIconX2
        End If
'        >> ).Grid.GColorDefault(1)
        If GridXYData(TmpCutGridX, TmpCutGridY).Grid.GColorDefault(1) = True Then _
            tmpColorGrid = GridXYData(TmpCutGridX, TmpCutGridY).Grid.FillColor Else _
            tmpColorGrid = SetNewGrid.GridXYFillColor
            
        ObjMe.Font.Bold = GridXYData(TmpCutGridX, TmpCutGridY).Grid.Bold
        ObjMe.Font.Italic = GridXYData(TmpCutGridX, TmpCutGridY).Grid.Italic
        ObjMe.Font.Underline = GridXYData(TmpCutGridX, TmpCutGridY).Grid.Underline
'        Picture1.FillColor = vbRed
'        Picture1.Refresh
        
        TextEffect ObjMe.hdc, GridXYData(TmpCutGridX, TmpCutGridY).GridXYValue, _
        GridX(TmpCutGridX).GridLeft + TmpRangePicGridX1 + TmpRangePicGridX2 + 5, GridY(TmpCutGridY).GridTop + 3, GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 5 - OverRangeEndHeads, GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight, , GridXYData(TmpCutGridX, TmpCutGridY).Grid.Alignment, tmpColorGrid
    End If
    
    If GridY(TmpCutGridY).GHeightDefault = True Then
    ObjMe.Line (GridX(TmpCutGridX).GridLeft, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.SellHeight_Def)-(Cx - OverRangeEndHeads, GridY(TmpCutGridY).GridTop + 25), 0, B
    
    If Pilih = 0 Or Pilih = 2 Then
            DefText ObjMe
            ObjMe.Font.Bold = GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Bold
            ObjMe.Font.Italic = GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Italic
            ObjMe.Font.Underline = GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Underline
            
            If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.TypeControl = "0" Then
                If GridXYData(TmpCutGridX, TmpCutGridY).GridSub.GColorDefault(0) = True Then _
                    tmpColorGrid = GridXYData(TmpCutGridX, TmpCutGridY).GridSub.BackColor Else _
                    tmpColorGrid = SetNewGrid.GridXYBackColorSub
                    
                If GridY(TmpCutGridY).GHeightDefault = True Then _
                    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 0, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.SellHeight_Def + 1)- _
                              (Cx - OverRangeEndHeads, Cy), tmpColorGrid, BF
                
                If GridXYData(TmpCutGridX, TmpCutGridY).GridSub.GColorDefault(1) = True Then _
                    tmpColorGrid = GridXYData(TmpCutGridX, TmpCutGridY).GridSub.FillColor Else _
                    tmpColorGrid = SetNewGrid.GridXYFillColorSub
                
                TextEffect ObjMe.hdc, GridXYData(TmpCutGridX, TmpCutGridY).GridXYValueSub, _
                GridX(TmpCutGridX).GridLeft + 5, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSize.SellHeight_Def + 1, GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 5 - OverRangeEndHeads, GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight, , GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Alignment, tmpColorGrid
            End If
        End If
    End If

'repair text and sub text

    ObjMe.Line (GridX(TmpCutGridX).GridLeft, GridY(TmpCutGridY).GridTop)-(Cx - OverRangeEndHeads + 0, Cy), 0, B
    If OnlyLineHead = False And OverRangeEndHeads > 0 And HeadCountAuto = False Then
        TmpColorsLineHead = LinePosisionHeads()
        DrawingLineHeadCol ObjMe, _
        RangeEndHead, 5, TmpColorsLineHead, Val(TmpCutGridY)
        
'        DrawingLineHeadCol ObjMe, _
        GridX(HeadSellOnFixCol(3 - 1 + HscHead.Value)).GridLeft + GridX(HeadSellOnFixCol(3 - 1 + HscHead.Value)).GridWidth, 5, vbRed 'TmpColorsLineHead
    End If
Case 1 'Y
    If GridY(TmpCutGridY).GridStyle = 1 Then 'One color
        ObjMe.Line (0, GridY(TmpCutGridY).GridTop)-(SetNewGrid.GridSize.GDRangeX, Cy), GridY(TmpCutGridY).GridColGra(0), BF
    ElseIf (GridY(TmpCutGridY).GridStyle = 2 Or GridY(TmpCutGridY).GridStyle = 3) Then 'Two color
        pvDrawGrid ObjMe, 1, Val(GridY(TmpCutGridY).GridTop + 1), Val(SetNewGrid.GridSize.GDRangeX), Val(Cy), GridY(TmpCutGridY).GridColGra(0), GridY(TmpCutGridY).GridColGra(1), GridY(TmpCutGridY).GridStyle - 2
    ElseIf GridY(TmpCutGridY).GridStyle = 4 Then 'Picture
        
    End If
    
    ObjMe.Line (1, GridY(TmpCutGridY).GridTop + 1)-(SetNewGrid.GridSize.GDRangeX + 0, Cy), vbWhite, B
    ObjMe.Line (0, GridY(TmpCutGridY).GridTop)-(SetNewGrid.GridSize.GDRangeX + 0, Cy), , B
    If GridY(TmpCutGridY).GridHeight <= SetNewGrid.GridSize.SellHeight_Def Then
        DrawSubY ObjMe, 0, 0, TmpCutGridY
        f = 1
    Else
        DrawSubY ObjMe, 1, 0, TmpCutGridY
        If GridY(TmpCutGridY).GridHeight > SetNewGrid.GridSize.SellHeight_Def + (SetNewGrid.GridSize.SellHeight_Def / 4) Then _
        f = 2 Else f = 1
    End If
    
    cf = (GridY(TmpCutGridY).GridTop + ((GridY(TmpCutGridY).GridHeight - ObjMe.TextHeight(GridY(TmpCutGridY).GridValue)) \ f))
'    D = GridY(TmpCutGridY).GridValue 'SetNewGrid.GridSize.SellHeight_Def  ' GridY(TmpCutGridY).GridValue
'    TextEffect ObjMe.hdc, GridY(TmpCutGridY).GridValue, _
    (SetNewGrid.GridSize.GDRangeX - ObjMe.TextWidth(D)) / 2, cf, , , 0
    TextEffect ObjMe.hdc, GridY(TmpCutGridY).GridValue, _
    0, cf, SetNewGrid.GridSize.GDRangeX, GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight, , 2, SetNewGrid.GridLenkapY.GridForeColor

Case 2 'X
    'ObjMe.Line (GridX(TmpCutGridX).GridLeft, 0)-(Cx, SetNewGrid.GridSize.GDRangeY - 1), &H8000000F, BF
    'bx9
    'If SetNewGrid.GridStyleX = 2 And GridX(TmpCutGridX).GridStyle = 0 Then
    '    pvDrawGrid ObjMe, Val(GridX(TmpCutGridX).GridLeft + 1), 1, Val(Cx), Val(SetNewGrid.GridSize.GDRangeY), vbBlue, vbRed, 0
    'End If
    
    If GridX(TmpCutGridX).GridStyle = 1 Then 'One color
        ObjMe.Line (GridX(TmpCutGridX).GridLeft, 0)-(Cx, SetNewGrid.GridSize.GDRangeY), GridX(TmpCutGridX).GridColGra(0), BF
    ElseIf GridX(TmpCutGridX).GridStyle = 2 Or GridX(TmpCutGridX).GridStyle = 3 Then 'Two color
        pvDrawGrid ObjMe, Val(GridX(TmpCutGridX).GridLeft + 1), 1, Val(Cx), Val(SetNewGrid.GridSize.GDRangeY), GridX(TmpCutGridX).GridColGra(0), GridX(TmpCutGridX).GridColGra(1), GridX(TmpCutGridX).GridStyle - 2
    ElseIf GridX(TmpCutGridX).GridStyle = 4 Then 'Picture
        
    End If
    'pvDrawGrid Val(GridX(TmpCutGridX).GridLeft + 1), 1, Val(Cx), Val(SetNewGrid.GridSize.GDRangeY), vbBlue, vbRed, 0
    
    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 1, 1)-(Cx, SetNewGrid.GridSize.GDRangeY - 0), vbWhite, B
    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 0, 0)-(Cx, SetNewGrid.GridSize.GDRangeY - 0), , B
    
'    Picture5.CurrentX = 0
'    Picture5.CurrentY = 0
'    Picture5.Print "OOOOOOOOO"
'    Picture5.Line (0, 0)-(100, 100), vbRed
    'cf = (GridX(TmpCutGridX).GridLeft + ((GridX(TmpCutGridX).GridWidth - ObjMe.TextWidth(GridX(TmpCutGridX).GridValue)) / 2))
'    TextEffect ObjMe.hdc, GridX(TmpCutGridX).GridValue, _
    cf, 2, , , 0
    

    ObjMe.Font.Bold = True
    'ObjMe.Font = "12121212"
    TextEffect ObjMe.hdc, GridX(TmpCutGridX).GridValue, _
    GridX(TmpCutGridX).GridLeft, 2, GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth, GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight, , 2, SetNewGrid.GridLenkapX.GridForeColor

End Select
End Sub

Private Sub test23(LIndex As String)
LIndex = 5
End Sub

Private Sub DrawButton(TmpCutGridX As Integer, TmpCutGridY As Integer, ObjMe As Object)
ObjMe.Line (0, 0)-(100, 100)
End Sub

'Private Sub DrawList(ByVal LIndex As Integer, TmpCutGridX As Integer, TmpCutGridY As Integer, ObjMe As Object, SizeGridList As Integer, CountGridList As Integer)
Private Sub DrawList(TmpCutGridX As Integer, TmpCutGridY As Integer, ObjMe As Object, Optional IndexLList As Integer, Optional MouseDowns As Boolean, Optional CloseBar As Boolean, Optional SubCloseBar As Boolean, Optional OnOutControl As Boolean)
Dim tmpSizeGridList As Integer
Dim XRangeList As Integer
Dim UIndexs As Integer, LIndex As Integer
Dim CountGridList As Integer, SizeGridList As Integer
Dim Y As Single
Dim OverLos As Integer
'Dim ListPT.PTX2 As Integer
Dim FormatD As String
Dim YyY As Integer
Dim nY As Single
Dim ColorHead As Long

'Exit Sub
'm s d 'cek untuk background list mengalami perubahan semuanya

'GridXYData(XPointerIndex, YPointerIndex).GridSubType.CountContGList

'^^^^
'            SizeGridList = 16
'            CountGridList = Int((GridY(TmpCutGridY).GridHeight - SellHeight_Def) / SizeGridList) + 0
Cx = GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 0
Cy = GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight

 
''GridXYData(3, 100).GridSubType.GLFrmtStly = "?pt.vis|1 ?pt.x2|150?"
''GridXYData(0, 150).GridSubType.GLFrmtStly = "" '?pt.vis|1 ?pt.x2|150 ?pt.full|1 ?" '?pt.x2|50 ?pt.full|0 ?pl.txt.x1|10 ?"
''GridXYData(2, 100).GridSubType.GLFrmtStly = "?pt.vis|1 ?pt.x2|150 ?lst.txt.x1|10 ?"
''GridXYData(0, 150).GridSubType.GLFrmtStly = "?pt.vis|1 ?pt.x2|150 ?pt.ico.file|D:\pic.bmp ?pt.index|1 ?"
'GridXYData(0, 150).GridSubType.GLShowCount = 3
'GridXYData(2, 100).GridSubType.GLFrmtStly = "?name|PicThumb ?.x2|25 ??name|PicList"
'" l?pt.x2|25 ?name|PicList ?pl.txt.x1|10 ?"
'MsgBox GridXYData(2, 100).GridSubType.GLFrmtStly
'MsgBox Get_Format(GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLFrmtStly, "?pt.full")
'MsgBox Get_Format("?" & Get_Format("?pt.full|0|5 ?", "?pt.full") & "?", "?0")

' gfdf gdfg dfg df gdf gd gd gd fgd SizeGridList dg dfg dg
FormatD = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLFrmtStly
LoadFormatPT FormatD
If ListPT.PTVis = 0 Then ListPT.PTFull = 0
If ListPT.PTX2 < 15 Then ListPT.PTVis = 0

If ListPT.PTIcon.File <> "" Then
    If Dir(ListPT.PTIcon.File) <> "" Then
        tmpImage.Picture = LoadPicture(ListPT.PTIcon.File)
    Else
        tmpImage.Picture = Nothing
    End If
Else
    tmpImage.Picture = Nothing
End If

'MsgBox ListPT.PTFull
'If TmpCutGridX = 0 Then
Command12.Caption = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLScrolIndex & " | " & LIndex & " - " & (GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLCount - 0) - LIndex & " " & CountGridList & " = " & CountGridList - ((GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLCount - 0) - LIndex)


'End If

'If LIndex < 0 Then LIndex = 0 ': Exit Sub

'ObjMe.Line (GridX(TmpCutGridX).GridLeft + 2, _
(GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLPointer + 0)))- _
(Cx - XRangeList, (GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLPointer + 1))), vbBlue, BF

Command8.Caption = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLPointer
Command8.Caption = GridY(TmpCutGridY).GridHeight - SellHeight_Def & " OPPP"

'mode auto : pertimbangkan buat mode manual
Dim tmpCountContGList As Integer, tmpGLRange As Integer
Dim CountList As Integer, LoopCountList As Integer, LoopLine As Integer
Dim tmpGLShowCount As Integer

'GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLShowCount = 0

'List1.Clear
'GridXYData(1, 4).GridSubType.GLShowCount = 2
tmpCountContGList = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1
tmpGLShowCount = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLShowCount - 1
'tmpGLShowCount = 1
'MsgBox tmpCountContGList
If tmpGLShowCount > -1 Then
    tmpCountContGList = tmpGLShowCount
        If tmpGLShowCount > GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1 Then
            tmpCountContGList = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1
            tmpGLShowCount = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1
        End If
End If
     
'--------------------------------
Dim Aab As Integer
Aab = 20 'ggggggg hhhhhhhhh
    
If tmpCountContGList = 0 Then CountList = GridY(TmpCutGridY).GridHeight - SellHeight_Def - 0 Else _
    CountList = ((GridY(TmpCutGridY).GridHeight - SellHeight_Def) \ (tmpCountContGList + 1))
    If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight <> 0 Then CountList = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight
'^ CUT---------------------------

'GridXYData(0, 100).GridSubType.GLSubScrolIndex = 1
List3.Clear
UIndexs = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLSubScrolIndex

'ObjMe.Line (GridX(TmpCutGridX).GridLeft + 1, GridY(TmpCutGridY).GridTop + SellHeight_Def)-(Cx, Cy), &H808080, BF










'MsgBox CountList & " " & TmpCutGridX
'Dim Picr As StdPicture


If ListPT.PTVis = 1 And CloseBar = False Then
'* Tambahkan If untuk tidak keperluan
'    MsgBox Get_Format(GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLFrmtStly, "?name")
'    MsgBox Get_Format(GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLFrmtStly, "?x2")
    'ListPT.PTX2 = Val(Get_Format(FormatD, "?pt.x2"))
    If ListPT.PTFull = 1 Then ListPT.PTX2 = GridX(TmpCutGridX).GridWidth
    
    Picture8.Picture = Nothing
    'If ListPT.PTX2 <= 0 Then ListPT.PTX2 = 50 '>>>> errr
    Picture8.Width = ListPT.PTX2 - 3
    Picture8.Height = CountList - 2
    
    'Picture8.PaintPicture Picture7.Picture, 7, 7, ListPT.PTX2 - 3 - 14, CountList - 2 - 14, 7 * 1 + 1, 7 * 1 + 1, 6, 6
'    errrrrrrrrrrr errrrrrrrrrrrrrrr errrrrrrrrrrrrrrrrrrr
    Picture8.PaintPicture Picture7.Picture, 0, 0, ListPT.PTX2 - 3, CountList - 2, 7 * 1 + 1, 7 * 1 + 1, 6, 6
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
''    Dim Aab As Integer
''    Aab = 20 'ggggggg hhhhhhhhh
''            penambahan list untuk manual, mainkan count pada list
    
    SizeGridList = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLRange
    LIndex = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLScrolIndex

''    If tmpCountContGList = 0 Then CountList = GridY(TmpCutGridY).GridHeight - SellHeight_Def - 0 Else _
''       CountList = ((GridY(TmpCutGridY).GridHeight - SellHeight_Def) \ (tmpCountContGList + 1))
    If SizeGridList < 1 Then Exit Sub
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
        'GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLScrolIndex = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLScrolIndex - ErrorX
    End If

    'Update
    If LIndex = -1 Then
        LIndex = 0
        GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(X + UIndexs).GLScrolIndex = LIndex
    End If


'List1.AddItem GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 1 & " " & CountGridList & " " & CountList Mod SizeGridList
    If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 1 < CountGridList - 0 Then
        If tmpGLShowCount = -1 Or tmpGLShowCount = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1 Then
            XRangeList = 4
        Else
            XRangeList = 15 + 4
        End If
               'tmpGLShowCount
    End If
    
    tmpGLRange = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLRange
    
    
    LoopCountList = (CountList * (X + 1)) + ((GridY(TmpCutGridY).GridTop + SellHeight_Def))
    LoopLine = (Abs(LoopCountList - (GridY(TmpCutGridY).GridTop + SellHeight_Def + CountList))) + Aab ' 20




List3.AddItem CountList - Aab

'On Over Back
ObjMe.Line _
(GridX(TmpCutGridX).GridLeft + 2, GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + 2)- _
(Cx - 2, GridY(TmpCutGridY).GridTop + LoopLine + CountList + 4), _
GridXYData(TmpCutGridX, TmpCutGridY).GridSub.BackColor, BF

'ObjMe.Line _
(GridX(TmpCutGridX).GridLeft + 2, GridY(TmpCutGridY).GridTop + SellHeight_Def + (CountList * IndexList) + 2)- _
(Cx - 2, GridY(TmpCutGridY).GridTop + SellHeight_Def + (CountList * IndexList) + CountList - YyY), _
GridXYData(TmpCutGridX, TmpCutGridY).GridSub.BackColor, BF

'List3.AddItem tmpCountContGList

'        MsgBox LoopLine fgffghgfh

'List1.AddItem LoopLine & " " & CountGridList
    'di sini bisa di pasang label
    
    'TextEffect ObjMe.hDC, GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X).GLCaption, _
    GridX(TmpCutGridX).GridLeft + 2, _
    (CountList * (X + 0)) + ((GridY(TmpCutGridY).GridTop + SellHeight_Def)) + 3, _
    GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 2, _
    (CountList * (X + 0)) + ((GridY(TmpCutGridY).GridTop + SellHeight_Def)) + 16, , GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Alignment, tmpColorGrid
'GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X).GLCaption
    
            'List1.AddItem X '& "  " & CountList & " " & CountGridList

    'If Y + LIndex >= GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList.GLCount - 0 Then Exit For 'Exit Sub
    'If Y + LIndex >= GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X).GLCount - 0 Then Exit For 'Exit Sub
    ''If tmpPointerListHead = X Then
    ''    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 2, _
    ''    ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 14)- _
    ''    (GridX(TmpCutGridX).GridLeft + GridX(TmpCutGridX).GridWidth - 2, (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 0)), vbWhite - 255, BF
    ''End If
    
    If tmpGLShowCount > -1 Then
'        tmpCountContGList = tmpGLShowCount
        aXRangeList = 1
    Else
        aXRangeList = 0
    End If
    
    If ListPT.PTVis = 1 Then 'Format PicThumb
'    ListPT.PTX2 = 100
        If ListPT.PTFull = 0 And GridX(TmpCutGridX).GridLeft + ListPT.PTX2 + 10 >= Cx - XRangeList + 2 Then
    '        MsgBox ""
            TMPx2_PicThumb = 2
            GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).OpenThumb = False
        Else
            'Line Thumb
            '''ObjMe.Line (GridX(TmpCutGridX).GridLeft + 2, _
            (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17)- _
            (GridX(TmpCutGridX).GridLeft + ListPT.PTX2 - 2, LoopCountList), vbRed, BF
            
'            ObjMe.PaintPicture Picture8.Picture, GridX(TmpCutGridX).GridLeft + 2, _
            (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17, _
            , , , 7 * 1 + 1, 7 * 1 + 1, 6, 6
            'Picture Thumb
            If gh = 7899 Then
                gh = 0 'dellllllll
            End If
'            errrrrrrrrrrr errrrrrrrrrrrrrrr errrrrrrrrrrrrrrrrrrr
            If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17) + CountList > Cy Then
                YyY = Cy - ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17) '((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17)
            Else
                YyY = CountList
            End If
            If YyY > 0 And CloseBar = False Then
                ObjMe.PaintPicture Picture8.Picture, GridX(TmpCutGridX).GridLeft + 2, _
                (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17, _
                , , , , , YyY
            End If
            GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).OpenThumb = True
            
'        ObjMe.PaintPicture Picture7.Picture, GridX(TmpCutGridX).GridLeft + 2, _
        (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine - 17, _
        ListPT.PTX2 - 3, CountList - 2, 7 * 1 + 1, 7 * 1 + 1, 6, 6

'        ObjMe.PaintPicture Image3.Picture, Cx - XRangeList + 2, _
        GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine, , CountList - Aab, 14 * 1, , 14, 13
            TMPx2_PicThumb = ListPT.PTX2
        End If
    Else
        TMPx2_PicThumb = 2
    End If

'    ObjMe.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, _
    (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine)- _
    (Cx - XRangeList + 2, YyY - 0), 0, B 'vbred
    
    
    
    If ListPT.PTFull = 0 Then  'Update
        If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.PointerListHead = X + UIndexs Then '&H00FF8080&
            Picture1.FontBold = True
            ColorHead = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLHeadColorClik
            'ColorHead = &HFFC0C0 'Color List Head Hit
        Else
            Picture1.FontBold = False
            ColorHead = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLHeadColor
            'ColorHead = vbGreen '&HFF8080 'Color List Head Default
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
''ObjMe.Line ((GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb) * 0, _
(((GridY(TmpCutGridY).GridTop + SellHeight_Def) * 0) + LoopLine) - 17)- _
(Cx - (15 * aXRangeList) - 2, ((((GridY(TmpCutGridY).GridTop + SellHeight_Def) * 0) + LoopLine) - 1) - YyY), ColorHead, BF
        
            'Text Head List
            TextEffect ObjMe.hdc, GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCaption & vbCrLf & "LLLLL", _
            GridX(TmpCutGridX).GridLeft + 5 + TMPx2_PicThumb, _
            ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 14, _
            Cx - (15 * aXRangeList) - 2, _
            ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - YyY, , GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Alignment, 0
        End If
    
    'Command17.Visible = True
    'Command17.Caption = SellHeight_Def + GridY(TmpCutGridY).GridTop + GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight * (X + 1) & " " & Cy
'        If SellHeight_Def + GridY(TmpCutGridY).GridTop + GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight * (X + 1) > Cy Then Exit For
'>>>>>>>>>>>>>>>>>>> Membuat auto
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
'            Command17.Caption = Rrrx '(((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine - Rrrx) - Cy & " 00000"
            If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine < Cy Then
            ObjMe.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, _
            ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine)- _
            (Cx - XRangeList + 2, ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine - Rrrx), &H8000000D, BF
            End If
            ''Picture1.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 17)-(Cx - XRangeList + 2, 100), vbRed
                                  '^ Kurang sempurna
''            ObjMe.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, _
            ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine)- _
            (Cx - XRangeList + 2, ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine - Rrrx), &H8000000D, BF


            tmpColorGrid = vbWhite
        Else
            tmpColorGrid = 0
        End If
        
        
        tmpSizeGridList = SizeGridList
            If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1))) + LoopLine > _
               LoopCountList Then
                tmpSizeGridList = Abs((CountList - Aab) - (CountGridList * SizeGridList)) 'LoopCountList
            'List1.AddItem tmpSizeGridList '& "  " & CountList & " " & CountGridList
            End If
     
    'Update
        If Y + LIndex >= GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 0 Then
            Exit For 'Exit Sub
        End If

If OverLos > Cy Then
'    MsgBox "Ok"
If ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine + tmpSizeGridList > Cy Then
tmpSizeGridList = (Cy) - (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine)  'GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + (CountGridList * SizeGridList) - Cy 'Abs((CountList - Aab) - (CountGridList * SizeGridList))
'Command17.Caption = (Cy) - (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine)   'GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + (CountGridList * SizeGridList) - Cy '((GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))) + LoopLine + tmpSizeGridList - Cy '(SizeGridList * (Y + 1))

''Picture1.Line (2 * (X + 1), 0)-(2 * (X + 1), GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + ((Y + 1) * SizeGridList)), (50 * (X + 1))
End If
End If
    Picture1.FontBold = False
    TextEffect ObjMe.hdc, GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLText(Y + LIndex), _
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

        
'ObjMe.Line (GridX(TmpCutGridX).GridLeft + 2, _
(GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0)) + ((0 + 0) * 0))- _
(Cx - XRangeList, (GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 1)) + ((0 + 0) * 0)), 0, B
        
'        revisi revisi 1111 222222223           3    3
        
    Next Y
    'List1.AddItem LoopLine

    'Update
    If Y + LIndex >= GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 0 Then
        'Exit For 'Exit Sub
    End If

'if GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLHeight <> 0
If OverLos > Cy Then YyY = Cy Else YyY = LoopCountList
If GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine < Cy Then 'MsgBox "O "
'Command17.Caption = GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + 0 & " " & Cy & " ."
    'Garis Back
    ObjMe.Line (GridX(TmpCutGridX).GridLeft + TMPx2_PicThumb, _
    (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine)- _
    (Cx - XRangeList + 2, YyY - 0), 0, B 'vbred
End If
    'Line Try
'    ObjMe.Line (GridX(TmpCutGridX).GridLeft + 2, LoopCountList)-(Cx - XRangeList + 2, LoopCountList), vbWhite
    'ObjMe.Line (GridX(TmpCutGridX).GridLeft + 2, LoopCountList + 1)-(Cx - XRangeList, LoopCountList + 1), 0
    

'-------[Picture Scroll]------------
    If GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCount - 1 > CountGridList - 1 Then
        'ObjMe.Line (Cx - XRangeList + 15, _
        (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine)- _
        (Cx - XRangeList + 2, LoopCountList), 0, B
        
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
        'ObjMe.Line (Cx - 15, _ kkkkk
        '(GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine)- _
        (Cx - 2, (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine + 12), vbBlue, B

        'ObjMe.Line (Cx - XRangeList + 15, _
        (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine)- _
        (Cx - XRangeList + 2, (GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine + 12), vbBlue, B
        'Kotak atas
'            Command17.Caption =
        If GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine + 13 > Cy Then YyY = (CountList - Aab) - (LoopCountList - Cy) Else YyY = 13
        If YyY > 0 Then
            ObjMe.PaintPicture Image3.Picture, Cx - XRangeList + 2, _
            GridY(TmpCutGridY).GridTop + SellHeight_Def + LoopLine, , , 14 * 0, , 14, YyY
        End If
        'ObjMe.Line (Cx - XRangeList + 15, LoopCountList - 12)-(Cx - XRangeList + 2, LoopCountList), vbBlue, B
'        ObjMe.PaintPicture Image3.Picture, Cx - XRangeList + 2, LoopCountList - 12, , , 14 * 2, , 14, 13
        'Kotak bawah
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
'ObjMe.Line (GridX(TmpCutGridX).GridLeft + 2, (GridY(TmpCutGridY).GridTop + SellHeight_Def))- _
(Cx - XRangeList, (GridY(TmpCutGridY).GridTop + GridY(TmpCutGridY).GridHeight) - 0), 0, B
    'Tengah

    'ObjMe.Line (Cx - 15, _
    (GridY(TmpCutGridY).GridTop + SellHeight_Def))- _
    (Cx - 2, Cy), 0, B
    ObjMe.PaintPicture Image3.Picture, Cx - 15, GridY(TmpCutGridY).GridTop + SellHeight_Def + 1, , GridY(TmpCutGridY).GridHeight - SellHeight_Def - 1, 14 * 1, 13 * 1, 14, 13
    

    
    'MsgBox X ' & " " & LoopLine
'MsgBox "ok"
    If SubCloseBar = False Then
        SubMovingYs = (GridY(TmpCutGridY).GridHeight - SellHeight_Def - (13 * 3)) / _
        ((GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1) - tmpGLShowCount - 0)
        
        'MsgBox ((GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.CountContGList - 1) - tmpGLShowCount)
        
        ObjMe.PaintPicture Image3.Picture, Cx - 15, _
        (GridY(TmpCutGridY).GridTop + SellHeight_Def + 0 + 13) + _
        (SubMovingYs * UIndexs), , , 14 * 3, 13 * 1, 14, 13
'        Picture1.PaintPicture
        List1.AddItem "OJ"
    'Dalam Tengah
        'MovingY = (GridY(TmpCutGridY).GridHeight - SellHeight_Def - (13 * 3)) / _
        (GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList.GLCount - CountGridList)
        'ObjMe.PaintPicture Image3.Picture, Cx - 15, _
        (GridY(TmpCutGridY).GridTop + SellHeight_Def + 13) + _
        (MovingY * GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList.GLScrolIndex), , , 14 * 3, , 14, 13
    End If
    
    'Kotak atas
    'ObjMe.Line (Cx - 15, _
    (GridY(TmpCutGridY).GridTop + SellHeight_Def))- _
    (Cx - 2, GridY(TmpCutGridY).GridTop + SellHeight_Def + 12), 0, B
    ObjMe.PaintPicture Image3.Picture, Cx - 15, GridY(TmpCutGridY).GridTop + SellHeight_Def + 1, , , 14 * 0, , 14, 13
    
    'Kotak bawah
    'ObjMe.Line (Cx - 15, (Cy - 13))-(Cx - 2, Cy - 1), 0, B
    ObjMe.PaintPicture Image3.Picture, Cx - 15, Cy - 13, , , 14 * 2, , 14, 13
End If

End Sub

Private Sub DrawPicControl(ObjMe As Object, TmpCutGridX As Integer, TmpCutGridY As Integer, Y As Single, LIndex As Integer, SizeGridList As Integer, Optional LastX1 As Single)
Dim Ghi As String, ax As Integer
Dim tmpPicture As StdPicture

'SetNewGrid.GridSizePic.SellIconY1 = (SetNewGrid.GridSize.SellHeight_Def - SetNewGrid.GridSizePic.SellIconY2) / 2

Ghi = Replace(GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(IndexGL).GLControl(Y + LIndex), " ", "")
If Ghi <> "" Then
    
    IndexY = Get_GLControl(Ghi, "?indx") \ (Get_GLControl(Ghi, "?Col") + 1)
    IndexX = Get_GLControl(Ghi, "?indx") - (IndexY * (Get_GLControl(Ghi, "?Col") + 1))

    '        MsgBox ghi
    Set tmpPicture = LoadPicture(Get_GLControl(Ghi, "?filn"))
    

    ax = Get_GLControl(Ghi, "?x2")
    'If SetNewGrid.GridSizePic.SellIconX2 + SetNewGrid.GridSizePic.SellIconX1 > GridX(TmpCutGridX).GridWidth Then _
        ax = GridX(TmpCutGridX).GridWidth - SetNewGrid.GridSizePic.SellIconX1 'MsgBox "O"
    ObjMe.PaintPicture tmpPicture, GridX(TmpCutGridX).GridLeft + Get_GLControl(Ghi, "?x1") + 3 + LastX1, _
    GridY(TmpCutGridY).GridTop + SellHeight_Def + (SizeGridList * (Y + 0)) + Get_GLControl(Ghi, "?y1") + 0, _
    , , ax * IndexX, _
    Get_GLControl(Ghi, "?y2") * IndexY, ax - 0, Get_GLControl(Ghi, "?y2") - 0
'        dfs sdf s f sf sf
        
        
    'Picture1.PaintPicture Image5.Picture, GridX(TmpCutGridX).GridLeft + 5, _
    (GridY(TmpCutGridY).GridTop + SellHeight_Def) + (SizeGridList * (Y + 0))
                'MsgBox Get_GLControl(Ghi, "?filn")

End If

End Sub



'Private Sub pvDrawGrid(F1 As RECT2, Color1 As OLE_COLOR, Color2 As OLE_COLOR, HonVen As Long)
Public Sub pvDrawGrid(ObjMe As Object, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color1 As OLE_COLOR, Color2 As OLE_COLOR, HonVen As Long)
Dim Col1 As RGB, Col2 As RGB, F1 As RECT2
    F1.X1 = X1
    F1.X2 = X2
    F1.Y1 = Y1
    F1.Y2 = Y2
        
    Col1 = pvGetColorRGB(pvGetColorLong(Color1))
    Col2 = pvGetColorRGB(pvGetColorLong(Color2))
    Call pvDrawBackGrad(ObjMe.hdc, F1, Col1, Col2, HonVen)
End Sub

Public Sub XXpvDrawGrid(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color1 As OLE_COLOR, Color2 As OLE_COLOR, HonVen As Long)
Dim Col1 As RGB, Col2 As RGB, F1 As RECT2
    F1.X1 = X1
    F1.X2 = X2
    F1.Y1 = Y1
    F1.Y2 = Y2
        
    Col1 = pvGetColorRGB(pvGetColorLong(Color1))
    Col2 = pvGetColorRGB(pvGetColorLong(Color2))
    Call pvDrawBackGrad(hdc, F1, Col1, Col2, HonVen)
End Sub

Private Function pvGetColorLong(Color As Long) As Long
    If (Color And &H80000000) Then
        pvGetColorLong = GetSysColor(Color And &H7FFFFFFF)
      Else
        pvGetColorLong = Color
    End If
End Function

Private Function pvGetColorRGB(Color As Long) As RGB

  Dim HexColor As String
        
    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    pvGetColorRGB.R = "&H" & Mid(HexColor, 5, 2) & "00"
    pvGetColorRGB.G = "&H" & Mid(HexColor, 3, 2) & "00"
    pvGetColorRGB.b = "&H" & Mid(HexColor, 1, 2) & "00"
End Function

Private Sub pvDrawBackGrad(ByVal hdc As Long, pRect As RECT2, Color1 As RGB, Color2 As RGB, ByVal Direction As Long)

  Dim uTV(1) As TRIVERTEX
  Dim uGR    As GRADIENT_RECT
    
    '-- from
    With uTV(0)
        .X = pRect.X1
        .Y = pRect.Y1
        .R = Color1.R
        .G = Color1.G
        .b = Color1.b
        .Alpha = 0
    End With
    '-- to
    With uTV(1)
        .X = pRect.X2
        .Y = pRect.Y2
        .R = Color2.R
        .G = Color2.G
        .b = Color2.b
        .Alpha = 0
    End With
    
    uGR.UpperLeft = 0
    uGR.LowerRight = 1

    Call GradientFillRect(hdc, uTV(0), 2, uGR, 1, Direction)
End Sub


Private Sub DefText(ObjMe As Object)
    ObjMe.Font.Bold = False
    ObjMe.Font.Italic = False
    ObjMe.Font.Underline = False
End Sub

Sub DrawSubY(ObjMe As Object, IndexX As Integer, IndexY As Integer, TmpCutGridY As Integer)
'If SetNewGrid.GridSizePicSub.RangePicSubX1 < 1 Then SetNewGrid.GridSizePicSub.RangePicSubX1 = 3
'If SetNewGrid.GridSizePicSub.RangePicSuby1 < 1 Then SetNewGrid.GridSizePicSub.RangePicSuby1 = 3
If SetNewGrid.GridSizePicSub.RangePicSubX2 < 1 Then SetNewGrid.GridSizePicSub.RangePicSubX2 = 9
If SetNewGrid.GridSizePicSub.RangePicSubY2 < 1 Then SetNewGrid.GridSizePicSub.RangePicSubY2 = 9

    ObjMe.PaintPicture PictureSubGridY, _
    SetNewGrid.GridSizePicSub.RangePicSubX1, GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePicSub.RangePicSubY1, , , SetNewGrid.GridSizePicSub.RangePicSubX2 * IndexX, SetNewGrid.GridSizePicSub.RangePicSubX2 * IndexY, SetNewGrid.GridSizePicSub.RangePicSubX2, SetNewGrid.GridSizePicSub.RangePicSubY2
End Sub

Sub SLinkPictureSub(ObjMe As Object)
Set PictureSub = ObjMe
End Sub

'Sub SLinkPictureXY(ObjMe As Object)
'    If SetNewGrid.GridFilePicture = "" Then PicErr = True
'    Set PictureXY = ObjMe
'End Sub

Sub Initial()
    SetNewGrid.GridXCount = 1
    SetNewGrid.GridYCount = 1
    ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
    ReDim Preserve GridY(SetNewGrid.GridYCount - 1)
    ReDim Preserve GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
End Sub

Sub Add(Optional ByVal VX As Long, Optional ByVal VY As Long, Optional NewType As TypeAdd = -1)
Dim TmpGridXYData() As TypeGridXYData
Dim TmpVX As Long, TmpVY As Long
Dim A As Boolean, CountAddShowFix As Integer ', TmPVX As Integer

    If VY < 0 Or VY < 0 Then Exit Sub
    If VX > 0 Then
        If NewType = Fixed_X Or NewType = Fixed_XY Or NewType = Fixed_Sell_XY Then ReDim GridX(VX - 1) Else _
           ReDim Preserve GridX(VX - 1)
        
        If NewType = Fixed_Sell Or NewType = Fixed_Sell_XY Then
            ReDim GridXYData(VX - 1, VY - 1)
        Else
            TmpGridXYData() = GridXYData()
            
            If VY = 0 Then VY = SetNewGrid.GridYCount
            ReDim GridXYData(VX - 1, VY - 1)
            
            For X = 0 To VX - 1
                If X >= SetNewGrid.GridXCount - 0 Then Exit For
                For Y = 0 To VY - 1
                    If Y >= SetNewGrid.GridYCount - 0 Then Exit For
                    GridXYData(X, Y) = TmpGridXYData(X, Y)
                Next Y
            Next X
        End If
        
        If 8 = 8 Then
        CountAddShowFix = (VX - SellCountColumn) + ConShowFixCol
        If SellCountColumn = 0 Then CountAddShowFix = CountAddShowFix - 1: ConShowFixCol = -1
        
        ReDim Preserve ShowFixCol(CountAddShowFix)
        TmpVX = VX
        For X = CountAddShowFix To ConShowFixCol + 1 Step -1
            TmpVX = TmpVX - 1
            ShowFixCol(X) = TmpVX
            GridX(X).GridRealOnPosisi = -1
        Next X
        ConShowFixCol = CountAddShowFix
        
        End If
        
        SetNewGrid.GridXCount = VX
        If VY > 0 Then
            A = True
            SetNewGrid.GridYCount = VY
            ReDim Preserve GridY(VY - 1)
        End If
    End If
    If VY > 0 And A = False Then
        If NewType = Fixed_Y Or NewType = Fixed_XY Or NewType = Fixed_Sell_XY Then ReDim GridY(VY - 1) Else _
           ReDim Preserve GridY(VY - 1)
        ReDim Preserve GridXYData(SetNewGrid.GridXCount - 1, VY - 1)
        SetNewGrid.GridYCount = VY
    End If

If 8 = 9 Then
ReDim Preserve ShowFixCol(SellCountColumn - 1)
For X = 0 To SellCountColumn - 1
    ShowFixCol(X) = X
Next X
End If

Command21.Caption = ""
For X = 0 To UBound(ShowFixCol())
    Command21.Caption = Command21.Caption & ShowFixCol(X) & ","
Next X
Command21.Caption = Command21.Caption & vbCrLf & UBound(ShowFixCol())


Hsc.Max = ConShowFixCol 'SellCountColumn - 1
Vsc.Max = SellCountRow - 1
End Sub

Sub ClearRow()
    SetNewGrid.GridYCount = 1
    ReDim Preserve GridY(SetNewGrid.GridYCount)
    ReDim Preserve GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
    Vsc.Max = SetNewGrid.GridYCount - 1
End Sub

Sub ClearColumn()
    SetNewGrid.GridXCount = 1
    ReDim Preserve GridX(SetNewGrid.GridXCount)
    ReDim GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
    Hsc.Max = SetNewGrid.GridYCount - 1
End Sub

Sub ClearGrid()
    SetNewGrid.GridYCount = 1
    SetNewGrid.GridXCount = 1
    ReDim GridX(SetNewGrid.GridXCount)
    ReDim GridY(SetNewGrid.GridYCount)
    ReDim GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
    Vsc.Max = SetNewGrid.GridYCount - 1
    Hsc.Max = SetNewGrid.GridYCount - 1
    XPointerIndex = 0
    YPointerIndex = 0
    Pxy.Px1 = 0
    Pxy.Px2 = 0
    Pxy.Py1 = 0
    Pxy.Py2 = 0
    TmpPxy.Px1 = 0
    TmpPxy.Px2 = 0
    TmpPxy.Py1 = 0
    TmpPxy.Py2 = 0

    ConShowFixCol = 0
End Sub

'this up date to remove
Private Sub Tmp__RemoveXY(Optional ByVal IndexX As Long = -1, Optional ByVal IndexY As Long = -1)
Dim TmpGridXYData() As TypeGridXYData
Dim OpX As Boolean, OpY As Boolean
Dim VpX As Integer, VpY As Integer

        TmpGridXYData() = GridXYData()
        If IndexX > -1 Then
            IndexX = IndexX + 1
            OpX = True
        Else
            IndexX = 0
        End If
        If IndexY > -1 Then
            IndexY = IndexY + 1
            OpY = True
        Else
            IndexY = 0
        End If
        
        ReDim GridXYData(SetNewGrid.GridXCount - (1 + Abs(Int(OpX))), SetNewGrid.GridYCount - (1 + Abs(Int(OpY))))
              'GridXYData() = TmpGridXYData()
        
        For X = IndexX To SetNewGrid.GridXCount - 1
                If OpX = True Then GridX(X - 1) = GridX(X)
            For Y = IndexY To SetNewGrid.GridYCount - 1
                If OpY = True And X = IndexX Then GridY(Y - 1) = GridY(Y)
                If OpX = True Then GridXYData(X - 1, Y) = TmpGridXYData(X, Y)
                If OpY = True Then GridXYData(X, Y - 1) = TmpGridXYData(X, Y)
            Next Y
        Next X
        If OpX = True Then
            SetNewGrid.GridXCount = SetNewGrid.GridXCount - 1
            ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
            
            Hsc.Max = SetNewGrid.GridXCount - 1
        End If
        If OpY = True Then
            SetNewGrid.GridYCount = SetNewGrid.GridYCount - 1
            ReDim Preserve GridY(SetNewGrid.GridYCount - 1)
            
            Vsc.Max = SetNewGrid.GridYCount - 1
        End If
    
    Command3.Caption = IndexX
End Sub

Sub Remove1(Optional ByVal IndexX As Long = -1, Optional ByVal IndexY As Long = -1)
Dim TmpGridXYData() As TypeGridXYData
'    If IndexX = 0 Then IndexX = SetNewGrid.GridXCount
'    err rrrr rrr rrr
'    If IndexX > 0 Then
'        For X = IndexX To SetNewGrid.GridXCount - 1
'            GridX(X - 1) = GridX(X)
'        Next X
    If IndexX > -1 And IndexX <> SetNewGrid.GridXCount - 1 Then
        IndexX = IndexX + 1
        TmpGridXYData() = GridXYData()
'        ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
        SetNewGrid.GridXCount = SetNewGrid.GridXCount - 1
        ReDim GridXYData(SetNewGrid.GridXCount - 1, SetNewGrid.GridYCount - 1)
        
        GridXYData() = TmpGridXYData() 'di pertimbangkan
        
        For X = IndexX To SetNewGrid.GridXCount - 0
            If X >= IndexX Then GridX(X - 1) = GridX(X)
            For Y = 0 To SetNewGrid.GridYCount - 1
                'If X >= IndexX Then
                    GridXYData(X - 1, Y) = TmpGridXYData(X, Y) '???????
                'Else
                '    GridXYData(X, Y) = TmpGridXYData(X, Y)
                'End If
            Next Y
        Next X
'        SetNewGrid.GridXCount = SetNewGrid.GridXCount - 1
        ReDim Preserve GridX(SetNewGrid.GridXCount - 1)
    Else
    End If
'    End If
    
    Command3.Caption = IndexX
    'MsgBox SetNewGrid.GridXCount - 1
'    If IndexY = 0 Then IndexY = SetNewGrid.GridYCount
'    If IndexY > 0 Then
        'For Y = IndexY To SetNewGrid.GridYCount - 1
        '    GridY(Y - 1) = GridY(Y)
        'Next Y
    If IndexY > -1 And IndexY <> SetNewGrid.GridYCount - 1 Then
        IndexY = IndexY + 1
        
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
    
    'Add -1, 0
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

Sub GoDrawing()
    GridXY Picture1, Hsc.Value, Vsc.Value ', , , , Texts
    Picture1.Picture = Picture1.Image
    DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On4    DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
End Sub







































'Control-------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AF_Open_Click()
    OpenMe
End Sub
Sub OpenMe()
If Dir("D:\Data" & Picture1.Tag & ".txt") <> "" Then
    Open "D:\Data" & Picture1.Tag & ".txt" For Binary As #1
    Get #1, , J
    Close #1
    SetDataX
    GridXY Picture1, Hsc.Value, Vsc.Value
End If
End Sub
Private Sub AF_Save_Click()
GetData Picture1.Tag
End Sub

Private Sub AF_SaveMe_Click()
    SaveMes
End Sub

Private Sub APX_AColumn_Click()
If CheckType(GridType, 0) = False Then
    Add , SellCountRow
    GridXY Picture1, Hsc.Value, Vsc.Value
    Picture1.Picture = Picture1.Image
    DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    Hsc.Max = SellCountColumn - 1
Else
    MsgBox "Data ini hanya dapat di baca saja"
Exit Sub
End If

'Me.Caption = SellCountColumn & " " & SellCountRow
End Sub

Private Sub APX_Icon_Click()
    IconShow XPointerIndex, -1
End Sub

Private Sub APX_RColumn_Click()
If CheckType(GridType, 0) = False Then
    Remove 0, -1
    GridXY Picture1, Hsc.Value, Vsc.Value
    Picture1.Picture = Picture1.Image
    DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    Hsc.Max = SellCountColumn - 1
Else
    MsgBox "Data ini hanya dapat di baca saja"
Exit Sub
End If
End Sub

Private Sub APXY_Icon_Click()
    IconShow XPointerIndex, YPointerIndex
End Sub

Private Sub APXY_Show_Click()
    Form4.Show
End Sub

Private Sub APY_ARow_Click()
If CheckType(GridType, 0) = False Then
    Add -1, 0
    GridXY Picture1, Hsc.Value, Vsc.Value
    Picture1.Picture = Picture1.Image
    'DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    Vsc.Max = SellCountRow - 1
    'Me.Caption = SellCountColumn & " " & SellCountRow
Else
    MsgBox "Data ini hanya dapat di baca saja"
Exit Sub
End If
End Sub

Private Sub APY_Icon_Click()
    IconShow -1, YPointerIndex
End Sub

Private Sub APY_RRow_Click()
If CheckType(GridType, 0) = False Then
    Remove -1, 0
    GridXY Picture1, Hsc.Value, Vsc.Value
    Picture1.Picture = Picture1.Image
    DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    Vsc.Max = SellCountRow - 1
Else
    MsgBox "Data ini hanya dapat di baca saja"
Exit Sub
End If
End Sub

Sub IconShow(IndexX As Long, IndexY As Long)
    If CheckType(GridType, 0) = False Then
        If IndexX = -1 Then Form5.Text1(0).SellText = "All" Else Form5.Text1(0).SellText = IndexX
        If IndexY = -1 Then Form5.Text1(1).SellText = "All" Else Form5.Text1(1).SellText = IndexY
        If IndexX = -1 Or IndexY = -1 Then Form5.Text1(2).SellText = 0 Else Form5.Text1(2).SellText = SellIconIndex(IndexX, IndexY)
        Form5.LoadData Me.Tag
        Form5.Show
        TmpPxy = Pxy
    Else
        Pic6Show "Data ini hanya dapat di baca saja"
        
        'MsgBox "Data ini hanya dapat di baca saja"
    End If
End Sub

Private Sub Command1_Click()
'Picture1.BackColor = UserControl.BackColor
MsgBox Picture1.Picture
Picture1.Picture = Nothing
MsgBox Picture1.Picture

'm = CheckType(GridType, 0)
'Me.Caption = m

'Me.nForm(cnForm - 1).SellIconPicContColms

'MsgBox Me.Tag
Exit Sub
'Add -1, 26
'For X = 0 To 25
'    SellText(0, X) = Chr(X + 65) '& " -----------"
'    SellText(1, X) = X + 1 'Chr(XY + 65) & " -----------"
'    SellText(2, X) = 0 'Chr(XY + 65) & " -----------"
'Next X
'SellIconIndex(2, 2) = 2

FixedWidth = 125
'nForm(cnForm - 1).FixedHeight = 20
GridXY Picture1, Hsc.Value, Vsc.Value
'Vsc.Max = SellCountRow - 1
End Sub

Private Sub Command11_Click()
GridXYData(1, 100).GridSubType.GLShowCount = 2
End Sub

Private Sub Command12_Click()
'Picture1.Line (0, 0)-(500, 500), , BF

'GridXY Picture1, Hsc.Value, Vsc.Value
DrawDrid Picture1, 0, 150, 0
Exit Sub
Static Addxx As Integer

Addxx = Addxx + 1
GridXYData(1, 0).GridSubType.ContGList(IndexGL).Add Addxx
GridXYData(1, 0).GridSubType.ContGList(IndexGL).GLText(Addxx) = Addxx & " BoedidaX9 "


GridXY Picture1, Hsc.Value, Vsc.Value
Picture1.Picture = Picture1.Image
End Sub

Private Sub Command18_Click()
GridXYData(0, 150).GridSubType.GLShowCount = 3
'GridXYData(0, 150).GridSubType.GLHeight = 50
'MsgBox UBound(GridXYData(0, 150).GridSubType.ContGList())
End Sub

Private Sub Command19_Click()
'GridXY Picture1, 0, Vsc.Value, , 2
GridXY Picture1, Hsc.Value, Vsc.Value, , 2
'DrawDrid Picture1, 5, 0, 0
End Sub

Private Sub Command2_Click()

Picture1.BackColor = 0
Exit Sub
Text1.Visible = False
Picture1.Cls
DrawInvert Picture1, SellLeft(XPointerIndex) + 3, SellTop(YPointerIndex) + 3, SellLeft(XPointerIndex) + SellWidth(XPointerIndex) - 3, SellTop(YPointerIndex) + SellHeight(YPointerIndex) - 3
Text1.Visible = True

End Sub

Private Sub Command21_Click()
Command21.Caption = ""
For X = 0 To UBound(ShowFixCol())
    Command21.Caption = Command21.Caption & ShowFixCol(X) & ","
Next X
Command21.Caption = Command21.Caption & vbCrLf & UBound(ShowFixCol())
End Sub

Private Sub Command26_Click()
    DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
End Sub

Private Sub Command3_Click()
Picture1.Cls
        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2

End Sub

Private Sub Command30_Click()
HeadOnHSC
End Sub

Private Sub Command5_Click()
Picture1.Cls
DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
End Sub

Private Sub Command8_Click()
'Timer1.Enabled = False
'MsgBox SellText(2, 0)
'Exit Sub

Add 99, 9999 ', -1   ', 2  ', 50
test test 12345

For X = 0 To SellCountColumn - 1
    FixedColsText(X) = "Label " & X
Next X

For X = 0 To SellCountColumn - 1
For Y = 0 To SellCountRow - 1
    SellIconIndex(X, Y) = 0
    SellText(X, Y) = X & "SellText " & Y + 1
    FixedRowsText(Y) = Y + 1 & "."
Next Y
Next X


Hsc.Max = SellCountColumn - 1
Vsc.Max = SellCountRow - 1
GridXY Picture1, Hsc.Value, Vsc.Value

Picture1.Picture = Picture1.Image


'Timer1.Enabled = True
End Sub

Private Sub Form_Activate()
'MsgBox Me.Tag
'Form5.Caption = Me.Caption
'Form5.Tag = Me.Tag
End Sub

Private Sub Form_Load()
'
End Sub

Private Sub Form_Resize()
On Error Resume Next
Picture1.SellWidth_Def = Me.ScaleWidth - Picture1.left * 2 - Vsc.SellWidth_Def
Picture1.SellHeight_Def = Me.ScaleHeight - Picture1.top * 2 - Hsc.SellHeight_Def
Picture2.SellWidth_Def = Picture1.ScaleWidth
Picture3.SellHeight_Def = Picture1.ScaleHeight
    Picture6.left = (Picture1.ScaleWidth - Picture6.SellWidth_Def) \ 2
    Picture6.top = (Picture1.ScaleHeight - Picture6.SellHeight_Def) \ 2

Vsc.left = Picture1.SellWidth_Def + Picture1.left
Vsc.top = Picture1.top
Vsc.SellHeight_Def = Picture1.SellHeight_Def
Hsc.left = Picture1.left
Hsc.top = Picture1.SellHeight_Def - Picture1.top + Hsc.SellHeight_Def
Hsc.SellWidth_Def = Picture1.SellWidth_Def

TableWidth = Me.ScaleWidth
TableHeight = Me.ScaleHeight
'
'Me.Show
'Hsc.Value = 22

GridXY Picture1, Hsc.Value, Vsc.Value
Picture1.Picture = Picture1.Image

Pxy.Px1 = 1: Pxy.Px2 = SellCountColumn - 1: Pxy.Py1 = 0: Pxy.Py2 = 5
Picture1.Cls
DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
DragXY = True
TypeDragMove = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Command3.Caption = ""
'K = C
'For XY = Check1.Count - 1 To 0 Step -1
'    If K Mod (XY + 2) = 0 Then
'        K = K / (XY + 2)
'        Check1(XY).Value = 0
'        Command3.Caption = Command3.Caption & XY & " "
'        FileAll(IndexOpenFileAll).GridNew.GridType
'    End If
'Next XY
'C = 1
If CheckType(GridType, 2) = True Then
    Cancel = 1
    Pic6Show "Data ini tidak dapat di close"
    Exit Sub
End If

Dim Msg As Integer
Msg = MsgBox("Apakah data tabel """ & NameTabel & """ akan di simpan", vbYesNoCancel)
    If Msg = vbYes Then
        If Picture1.Tag = "" Then
            ReDim Preserve FileAll(CountFileAll)
            ReDim Preserve XForm(CountFileAll)

            GetData CountFileAll
            Form4.AddList CountFileAll
            CountFileAll = CountFileAll + 1
        Else
            GetData Picture1.Tag
        End If
        If Picture1.Tag <> "" Then XForm(Picture1.Tag).OnOff = False
    ElseIf Msg = vbCancel Then
        Cancel = 1
        Exit Sub
    Else
        If Picture1.Tag <> "" Then XForm(Picture1.Tag).OnOff = False
    End If
Me.Hide
Cancel = 1
End Sub

Private Sub Command4_Click()
    SellSubText(XPointerIndex, YPointerIndex) = Text2.Text
    Text2.Visible = False
    Picture1.SetFocus
    Fx = False
    Fy = False
    GridXY Picture1, Hsc.Value, Vsc.Value
    Picture1.Picture = Picture1.Image
'On        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
        DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
End Sub

Private Sub picShow_Click()
picShow.Cls
DrawList 0, 0, picShow, , , , , True
End Sub

Private Sub Picture1_DblClick()

If CheckType(GridType, 0) = True Then Exit Sub
                

If OnClickSellIncon = True Then
    RaiseEvent ClikSellIcon(XPointerIndex, YPointerIndex)
    Exit Sub
End If

If GridX(FixdColnIndex(XPointerIndex)).GridLeft + GridX(FixdColnIndex(XPointerIndex)).GridWidth > Picture1.ScaleWidth And _
   GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight > Picture1.ScaleHeight Then
    ScrollHRL XPointerIndex, YPointerIndex, True, True
ElseIf GridX(FixdColnIndex(XPointerIndex)).GridLeft + GridX(FixdColnIndex(XPointerIndex)).GridWidth > Picture1.ScaleWidth Then
    ScrollHRL XPointerIndex, , True
ElseIf GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight > Picture1.ScaleHeight Then
    ScrollHRL , YPointerIndex, , True
End If

If (DblX > FixedWidth And DblY < FixedHeight) Then
    AllText Text1, FixedColsText(FixdColnIndex(XPointerIndex)), SellLeft(FixdColnIndex(XPointerIndex)), 0, SellWidth(FixdColnIndex(XPointerIndex)), FixedHeight
ElseIf (DblX < FixedWidth And DblY > FixedHeight) Then
    AllText Text1, FixedRowsText(YPointerIndex), 0, SellTop(YPointerIndex), FixedWidth, SellHeight(YPointerIndex)
Else
    If DblY > GridY(YPointerIndex).GridTop + SetNewGrid.GridSize.SellHeight_Def Then
        If GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.TypeControl <> 1 Then
            Text1.Visible = False
            AllText Text2, SellSubText(FixdColnIndex(XPointerIndex), YPointerIndex), SellLeft(FixdColnIndex(XPointerIndex)), SellTop(YPointerIndex) + SetNewGrid.GridSize.SellHeight_Def, SellWidth(FixdColnIndex(XPointerIndex)) + 18, Abs(GridY(YPointerIndex).GridHeight - SetNewGrid.GridSize.SellHeight_Def) + 18
            
            Command4.left = Text2.left + Text2.Width - 16
            Command4.top = Text2.top + Text2.Height - 16
            Command4.Visible = True
        End If
    Else
        Text2.Visible = False
        Command4.Visible = False
        AllText Text1, SellText(FixdColnIndex(XPointerIndex), YPointerIndex), SellLeft(FixdColnIndex(XPointerIndex)), SellTop(YPointerIndex), SellWidth(FixdColnIndex(XPointerIndex)), SellHeight_Def
    End If
End If


Command26.Caption = XPointerIndex
End Sub

Private Sub AllText(vObj As Object, vText As String, vLeft As Integer, vTop As Integer, vWidth As Integer, vHeight As Integer)
    vObj.Text = vText
    vObj.left = vLeft + 1
    vObj.Width = vWidth - 1
    vObj.top = vTop + 1
    vObj.Height = vHeight - 1
    vObj.Visible = True
    vObj.SetFocus
    vObj.BorderStyle = 1
    
End Sub

Private Sub Picture1_GotFocus()
If Text1.Visible = True Then Text1.SetFocus
'MsgBox ""
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim KeyCodeOver As Boolean

'MsgBox Pxy.Py1 & " " & Pxy.Py2
'Exit Sub
'MsgBox GridDownY
'ShowFixCol
TmpKeyCode = KeyCode
TypeDragMove = True
    Select Case KeyCode
    'Case vbKeyReturn, vbKeyEscape 'Allow escape to abort editing
        
    Case vbKeyF2
        
    Case vbKeySpace
    
    Case vbKeyA

    Case vbKeyUp
        If Pxy.Py1 <> 0 Then
            Pxy.Py1 = Pxy.Py1 - 1
            Pxy.Py2 = Pxy.Py2 - 1
                If Pxy.Py1 < Vsc.Value Then Vsc.Value = Vsc.Value - 1 'Timer1.Enabled = True: HTPxy.Py1 = 1
        End If
'                Command3.Caption = Hsc.Value & " " & Pxy.Py1 'GridRightX & " " & iPxy.Px1

    Case vbKeyDown
        If Pxy.Py2 <> SetNewGrid.GridYCount - 1 Then
            Pxy.Py1 = Pxy.Py1 + 1
            Pxy.Py2 = Pxy.Py2 + 1
                If Pxy.Py2 > GridDownY - 1 Then Vsc.Value = Vsc.Value + 1 'Timer1.Enabled = True: HTPxy.Py1 = 1
        End If
    Case vbKeyLeft
        If Pxy.Px1 <> 0 Then
            Pxy.Px1 = Pxy.Px1 - 1
            Pxy.Px2 = Pxy.Px2 - 1
                If Pxy.Px1 < Hsc.Value Then Hsc.Value = Hsc.Value - 1 'Timer1.Enabled = True: HTPxy.Py1 = 1
        End If
    Case vbKeyRight
        If Pxy.Px2 <> SetNewGrid.GridXCount - 1 - HideCountFixCol Then
            Pxy.Px1 = Pxy.Px1 + 1
            Pxy.Px2 = Pxy.Px2 + 1
                If Pxy.Px2 > GridRightX - 1 Then Hsc.Value = Hsc.Value + 1                    'Timer1.Enabled = True: HTPxy.Py1 = 1
        Else
            Exit Sub
        End If
    Case vbKeyPageUp
    
    Case vbKeyPageDown
    
    Case vbKeyHome
        
    Case vbKeyEnd
    
    Case Else
        KeyCodeOver = True
    End Select
    
    XPointerIndex = Pxy.Px1
    YPointerIndex = Pxy.Py1
    RaiseEvent KeyDown(XPointerIndex, YPointerIndex, KeyCode, Shift)
    
    If KeyCodeOver = True Then Exit Sub
    
    DoEvents
    Picture1.Cls
    GridXY Picture1, Hsc.Value, Vsc.Value
    DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On    DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
        
                    Command2.Caption = Vsc.Value & " O " & GridDownY & " " & GridRightX

        If GridX(ShowFixCol(Pxy.Px2)).GridLeft + GridX(ShowFixCol(Pxy.Px2)).GridWidth > Picture1.ScaleWidth Then
            TypeDragMove = False ' nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn
            Hsc.Value = SearchOverGrid(Pxy.Px2, False)   'XYZ + 1
                If GridX(ShowFixCol(Pxy.Px2)).GridLeft + GridX(ShowFixCol(Pxy.Px2)).GridWidth > Picture1.ScaleWidth Then Hsc.Value = Hsc.Value + 1 'pertimbangan
            TypeDragMove = True
        End If
        Command1.Caption = ""
        If GridY(Pxy.Py2).GridTop + GridY(Pxy.Py2).GridHeight > Picture1.ScaleHeight Then
            TypeDragMove = False
            Vsc.Value = SearchOverGrid(Pxy.Py2, True)
                If GridY(Pxy.Py2).GridTop + GridY(Pxy.Py2).GridHeight > Picture1.ScaleHeight Then Vsc.Value = Vsc.Value + 1 'pertimbangan
            TypeDragMove = True
        End If


'    Picture1.Cls
'    DrawInvertToGrid cbgrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
End Sub

Private Function SearchOverGrid(XYPointer As Long, XYOverGrid As Boolean) As Long
Dim XYOver As Long
    'ShowFixCol
    If XYOverGrid = False Then XYOver = SetNewGrid.GridSize.GDRangeX
    If XYOverGrid = True Then XYOver = SetNewGrid.GridSize.GDRangeY
    'Command21.Caption = Hsc.Value
    For XYZ = XYPointer + 1 To 0 Step -1
        If XYOverGrid = False Then
            If XYZ < SetNewGrid.GridXCount - HideCountFixCol Then
                XYOver = XYOver + GridX(ShowFixCol(XYZ)).GridWidth
                If XYOver > Picture1.ScaleWidth Then Exit For
            End If
        End If
        If XYOverGrid = True Then
            XYOver = XYOver + GridY(XYZ).GridHeight
            If XYOver >= Picture1.ScaleHeight Then Exit For
        End If
    Next XYZ
    SearchOverGrid = XYZ + 0
End Function

Private Sub Picture1_KeyPress(KeyAscii As Integer)
'
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
TypeDragMove = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim TmpPy2 As PointInvert

'If Text1.Visible = True Then Exit Sub

DblX = X
DblY = Y
Text1.Visible = False
Text2.Visible = False
Command4.Visible = False
'            Command3.Caption = ""
'test test 12345
'RaiseEvent

If Picture1.MousePointer = 2 Then
    If Button = 1 Then
        If X > FixedWidth And Y < FixedHeight Then
            Picture3.Visible = True
            Picture3.left = X
        End If
        If X < FixedWidth And Y > FixedHeight Then
            Picture2.Visible = True
            Picture2.top = Y
        End If
    End If
Else
    If ClickScroll = True Then 'Pengklikan scrol
        ClickScroll = False
        Picture1.Cls
        GridXY Picture1, Hsc.Value, Vsc.Value
        Picture1.Picture = Picture1.Image
        If CheckType(GridType, 0) = False Then _
            DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On6            DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    End If
    
    'MDown = True
    
        XPointerIndex = Hsc.Value
        YPointerIndex = Vsc.Value
        SearchPointer CBGrid, X, XPointerIndex, Y, YPointerIndex
         
        If XPointerIndex < 0 Then
            'XPointerIndex = HeadSellOnFixCol(Abs(XPointerIndex) - 1)
'            XPointerIndex = GridX(HeadSellOnFixCol(Abs(XPointerIndex) - 1)).GridRealOnPosisi
            XPointerIndex = GridX(HeadSellOnFixCol(Abs(XPointerIndex) - 1)).GridRealPosisi
            gh = gh
        End If
'    MsgBox CheckType(GridType, 0, 1)
'    If CheckType(GridType, 0) = False Then _
        Picture1.Cls '---> Test Invert

    
    'End If
    If (X > FixedWidth And Y < FixedHeight) Then       'X
        If CheckType(GridType, 0) = False Then
            Picture1.Cls
            DrawInvertToGrid CBGrid, Picture1, XPointerIndex, 0, XPointerIndex, SellCountRow - 1
''On7            DrawInvertToGrid CBGrid, Picture1, Pxy, XPointerIndex, 0, XPointerIndex, SellCountRow - 1
        End If
''        If Button = 2 Then PopupMenu APopUpX
        TypeDrag = "X"
        DragXY = False
    
        TTPxy = Pxy 'Tmp
        TPxy.Px1 = XPointerIndex - Pxy.Px1: TPxy.Px2 = Pxy.Px2 - XPointerIndex
        TPxy.Py1 = YPointerIndex - Pxy.Py1: TPxy.Py2 = Pxy.Py2 - YPointerIndex
    
    ElseIf (X < FixedWidth And Y > FixedHeight) Then   'Y
        If CheckType(GridType, 0) = False Then
            Picture1.Cls
            DrawInvertToGrid CBGrid, Picture1, 0, YPointerIndex, SellCountColumn - 1, YPointerIndex
''On8            DrawInvertToGrid CBGrid, Picture1, Pxy, 0, YPointerIndex, SellCountColumn - 1, YPointerIndex
        End If
        If Button = 2 Then MsgBox "PopupMenu APopUpY"
        TypeDrag = "Y"
        DragXY = False
        
        TTPxy = Pxy 'Tmp
        TPxy.Px1 = XPointerIndex - Pxy.Px1: TPxy.Px2 = Pxy.Px2 - XPointerIndex
        TPxy.Py1 = YPointerIndex - Pxy.Py1: TPxy.Py2 = Pxy.Py2 - YPointerIndex

    Else                                                    'XY
        
        'If GridY(YPointerIndex).GHeightDefault = True And GridXYData(XPointerIndex, YPointerIndex).GridSubType.TypeControl = 1 Then
        If Y > GridY(YPointerIndex).GridTop + SellHeight_Def Then 'ListBoX
            MouseDownList X, Y
            Picture1.Picture = Picture1.Image
'        hjghj j hg jhg jhgj
        Else
            If (XPointerIndex >= Pxy.Px1 And XPointerIndex <= Pxy.Px2) And (YPointerIndex >= Pxy.Py1 And YPointerIndex <= Pxy.Py2) Then
            'Definisi Drag Long
                DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On9                DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
                    DrawInvert Picture1, SellLeft(FixdColnIndex(XPointerIndex)) + 3, SellTop(YPointerIndex) + 3, SellLeft(FixdColnIndex(XPointerIndex)) + SellWidth(FixdColnIndex(XPointerIndex)) - 3, SellTop(YPointerIndex) + SellHeight(YPointerIndex) - 3
                DragXY = True
                TypeDrag = "[XY]"
                
                TTPxy = Pxy 'Tmp ShowFixCol
                TPxy.Px1 = XPointerIndex - Pxy.Px1: TPxy.Px2 = Pxy.Px2 - XPointerIndex
                TPxy.Py1 = YPointerIndex - Pxy.Py1: TPxy.Py2 = Pxy.Py2 - YPointerIndex
    '        Command4.Caption = Pxy.Px1 & " " & Pxy.Py1 & " " & Pxy.Px2 & " " & Pxy.Py2 '(Pxy.Px2 - Pxy.Px1) + 1
            Else
            'Definisi Drag Sort &&&
                TPxy.Px1 = 0: TPxy.Px2 = 0
                TPxy.Py1 = 0: TPxy.Py2 = 0
                
                Command25.Caption = Pxy.Px1
                If CheckType(GridType, 0) = False Then _
                    DrawInvertToGrid CBGrid, Picture1, XPointerIndex, YPointerIndex, XPointerIndex, YPointerIndex
''On10                    DrawInvertToGrid CBGrid, Picture1, Pxy, XPointerIndex, YPointerIndex, XPointerIndex, YPointerIndex
                
'                If Button = 2 Then PopupMenu APopUpXY
                DragXY = False
                TypeDrag = "XY"
                
                If SellIcon(XPointerIndex) = False Then RaiseEvent ClikSell(XPointerIndex, YPointerIndex)
            End If
        End If
        
'Me.Caption = SellLeft(Pxy.Px1) + SellWidth(Pxy.Px2) & " > " & Picture1.ScaleWidth    'Pxy.Px1 & " " & Pxy.Py1 & " " & Pxy.Px2 & " " & Pxy.Py2
'Command2.Caption = GridDownY   '(Pxy.Px2 - Pxy.Px1) + 1
'Hsc.Value = Hsc.Value + 1
'USE -------------------------------------
        OnClickSellIncon = False 'FixdColnIndex
'        If SellIcon(ShowFixCol(XPointerIndex)) = True Then
        If SellIcon(FixdColnIndex(XPointerIndex)) = True Then
            If (X > SellLeft(FixdColnIndex(XPointerIndex)) + SellIconX1 And _
            X < SellLeft(FixdColnIndex(XPointerIndex)) + SellIconX1 + SellIconX2) And _
            (Y > SellTop(YPointerIndex) + SellIconY1 And _
            Y < SellTop(YPointerIndex) + SellIconY1 + SellIconY2) Then
                OnClickSellIncon = True
                RaiseEvent ClikSellIcon(XPointerIndex, YPointerIndex)
                'ClickProses
                'MsgBox "OOO"
                'If Button = 1 Then _
                ClickProses Me.Tag, XPointerIndex, YPointerIndex 'bx9
            Else
                RaiseEvent ClikSell(XPointerIndex, YPointerIndex)
            End If
        End If
'-----------------------------------------
    End If
End If
'--------------------------------------------------------------------------------------------------------------------------------
           ' MouseDownList X, Y

End Sub

Private Sub MouseDownList(X As Single, Y As Single)
Dim Test1 As Integer, Test2 As Integer
    
Dim tmpCountContGList As Integer, tmpGLRange As Integer
Dim LoopCountList As Integer, LoopLine As Integer
Dim tmpGLShowCount As Integer, SizeGridList As Integer
Dim CountGridList As Integer
Dim Aab As Integer, tmpGLCount As Integer
Dim IndexListGrid As Integer
Dim tmpYList As Integer, YList As Integer
Dim rz As Integer
Dim YyY As Single

'Exit Sub
If GridY(YPointerIndex).GHeightDefault = True And GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.TypeControl = 1 Then
    
    Cx = GridX(ShowFixCol(XPointerIndex)).GridLeft + GridX(ShowFixCol(XPointerIndex)).GridWidth - 0
    Cy = GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight

List1.Clear

Aab = 20

'FormatD = GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.GLFrmtStly
'LoadFormatPT FormatD
'GridXYData(0, 150).GridSubType.GLHeight

LoadFormatPT GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLFrmtStly
If ListPT.PTVis = 0 Then ListPT.PTFull = 0
If ListPT.PTX2 < 15 Then ListPT.PTVis = 0


tmpGLShowCount = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLShowCount - 1
UIndexs = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex - 0
'CountList >> pertimbangkan nama variabel beda dengan isinya dan di compar k "sub DrawList"
tmpCountContGList = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.CountContGList - 1
    If tmpGLShowCount > -1 Then
        tmpCountContGList = tmpGLShowCount
        XRangeList = 2
    Else
        XRangeList = 1
    End If
If GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLHeight > 0 Then
    CountList = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLHeight
Else
    CountList = ((GridY(YPointerIndex).GridHeight - SellHeight_Def) \ (tmpCountContGList + 1))
End If
YList = (Y - (SellHeight_Def + GridY(YPointerIndex).GridTop))
IndexList = YList \ CountList

rz = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead - UIndexs - 0

    Command18.Width = 30
    Command18.Caption = IndexList
    
    ''SizeGridList = GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLRange
    ''tmpGLCount = GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLCount
    ''YListGrid = (Y - (SellHeight_Def + GridY(YPointerIndex).GridTop + Aab + (CountList * IndexList)))
    
    
    ''CountGridList = Int((CountList - Aab) \ SizeGridList) ' + 0
    
    ''If ListPT.PTVis = 1 Then
    ''    If GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).OpenThumb = False Then
    ''        TMPx2_PicThumb = 2
    ''    Else
    ''        TMPx2_PicThumb = ListPT.PTX2
    ''    End If
'''        If GridX(XPointerIndex).GridLeft + ListPT.PTX2 + 10 >= Cx - (15 * XRangeList) + 2 Then
'''            TMPx2_PicThumb = 2
'''            If tmpGLCount <= CountGridList Then TMPx2_PicThumb = ListPT.PTX2
'''        Else
'''            TMPx2_PicThumb = ListPT.PTX2
'''        End If
    ''Else
    ''    TMPx2_PicThumb = 2
    ''End If
    

'MsgBox CountList
    '46464 454646 654 64645 564
'GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(X + UIndexs).GLCount - 1 > CountGridList - 1
    'MsgBox CountGridList 'TMPx2_PicThumb '212
    
    If tmpGLShowCount > -1 And (X > Cx - 15 And X < Cx) Then
'    update nyari scrol sub
        If (Y > GridY(YPointerIndex).GridTop + SellHeight_Def And Y < GridY(YPointerIndex).GridTop + SellHeight_Def + 13) Then
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex - 1
                If GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex < 0 Then GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex = 0
        ElseIf (Y > (GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight) - 13 And Y < GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight + SellHeight_Def) Then
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex + 1

  Command18.Width = 300
  Command18.Caption = Cy \ (CountList + 0) '\ Cy  '((tmpGLShowCount + 1) * GridXYData(XPointerIndex, YPointerIndex).GridSubType.CountContGList & " 0000000"
            'if ASda > sd
            'GridXYData(XPointerIndex, YPointerIndex).GridSubType.CountContGList = 12
            With GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType
                If .GLSubScrolIndex > .CountContGList - (.GLShowCount - 1) - 1 Then _
                   .GLSubScrolIndex = .CountContGList - (.GLShowCount - 1) - 1
            End With
        Else
            'If (X > Cx - 15 And X < Cx) Then
                SubMovingYs = (GridY(YPointerIndex).GridHeight - SellHeight_Def - (13 * 3)) / _
                ((GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.CountContGList - 1) - tmpGLShowCount)
                
                ScrolY1 = (GridY(YPointerIndex).GridTop + SellHeight_Def + 13) + _
                (SubMovingYs * GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex)
                ScrolY2 = ScrolY1 + 13
                If Y > ScrolY1 And Y < ScrolY2 Then
                    
                    'MsgBox "OK"
                    TMP_GLSubScrolIndex = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex + 1
                    SubClikScrol = True ' MsgBox "o"
                    SubTmpScrolY = Y - ScrolY1 'Y
                    'DrawList test1, test2, Picture1, , , , True
'                        Picture1.Picture = Picture1.Image

                    'Command16.Caption = ScrolY1 & " " & Y & " = " & Y - ScrolY1
                End If
            'End If
        End If
        'UIndexs = GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex - 0
        If (Y > GridY(YPointerIndex).GridTop + SellHeight_Def And Y < GridY(YPointerIndex).GridTop + SellHeight_Def + 13) _
        Or (Y > (GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight) - 13 And Y < GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight + SellHeight_Def) Then
            Picture1.Line (GridX(ShowFixCol(XPointerIndex)).GridLeft + 2, (GridY(YPointerIndex).GridTop + SellHeight_Def + 1))- _
            (Cx - 2, (GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight) - 1), vbWhite, BF
        End If
        Test1 = ShowFixCol(XPointerIndex): Test2 = YPointerIndex
        DrawList Test1, Test2, Picture1 ', IndexList + 0, True
        
        Command17.Caption = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.CountContGList 'GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex
        Exit Sub
    End If
    
    
'=====================(untuk dibawah)==========================================================================================
    SizeGridList = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLRange
    tmpGLCount = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLCount
    YListGrid = (Y - (SellHeight_Def + GridY(YPointerIndex).GridTop + Aab + (CountList * IndexList)))
    
    If SizeGridList < 1 Then Exit Sub
    CountGridList = Int((CountList - Aab) \ SizeGridList) ' + 0
    
    'MsgBox CountGridList
    
    If ListPT.PTVis = 1 Then
        If GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).OpenThumb = False Then
            TMPx2_PicThumb = 2
        Else
            TMPx2_PicThumb = ListPT.PTX2
        End If
    Else
        TMPx2_PicThumb = 2
    End If
    
    If ListPT.PTFull = 0 And YListGrid >= 0 And tmpGLCount - 1 > CountGridList - 1 Then
        If (X > Cx - 15 * XRangeList And X < Cx - 15 * (XRangeList - 1)) And _
        (YListGrid > 0 And YListGrid < 13) Then 'Scroll Up
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex - 1
                GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
        ElseIf (X > Cx - 15 * XRangeList And X < Cx - 15 * (XRangeList - 1)) And _
        (YListGrid + Aab > CountList - 13) Then 'Scroll Down
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex = GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex + 1
                GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
        ElseIf YListGrid > 0 And (X > Cx - 15 * XRangeList And X < Cx - 15 * (XRangeList - 1)) Then  'Scroll Tengah
                MovingY = ((CountList - Aab + 1) - (13 * 3)) / _
                (GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLCount - (CountGridList + 0))
                List1.AddItem MovingY & " " & GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLCount
                
                GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
                    'If GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(rz + UIndexs).OpenThumb = False Then TMPx2_PicThumb = 2 Else TMPx2_PicThumb = ListPT.PTX2
                    
                    Picture1.Line (GridX(ShowFixCol(XPointerIndex)).GridLeft + TMPx2_PicThumb, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + 2)-(Cx - 15 * (XRangeList - 1) - 1, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + 19), vbWhite - 0, BF
                    Picture1.Line (GridX(ShowFixCol(XPointerIndex)).GridLeft + TMPx2_PicThumb, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab - 17)-(Cx - 15 * (XRangeList - 1) - 2, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab - 1), &HFF8080, BF
                    Picture1.FontBold = False
                    TextEffect Picture1.hdc, GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(rz + UIndexs).GLCaption, GridX(ShowFixCol(XPointerIndex)).GridLeft + 5 + TMPx2_PicThumb, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab - 14, Cx - 15 * (XRangeList - 1) - 2 + 0, _
                    GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab, , GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSub.Alignment, 0

                ScrolY1 = 13 + (MovingY * GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex)
                ScrolY2 = ScrolY1 + 13
                If YListGrid > ScrolY1 And YListGrid < ScrolY2 Then 'Scroll Barr
                    ClikScrol = True
                    TmpScrolY = YListGrid - ScrolY1
                    Exit Sub
                End If
        ElseIf X > GridX(ShowFixCol(XPointerIndex)).GridLeft + TMPx2_PicThumb And X < Cx - 15 * (XRangeList - 0) Then  'List Scroll Bar Sub
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLPointer = YListGrid \ SizeGridList + GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex
                GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
        Else 'Thumb
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
        End If
    ElseIf ListPT.PTFull = 0 And X > GridX(ShowFixCol(XPointerIndex)).GridLeft + TMPx2_PicThumb And YListGrid >= 0 Then   'List No Scroll Bar Sub
        GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLPointer = YListGrid \ SizeGridList + GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
    ElseIf ListPT.PTFull = 0 And X > GridX(ShowFixCol(XPointerIndex)).GridLeft + TMPx2_PicThumb And YListGrid < 0 Then  'List Head Scroll Bar
        If tmpGLShowCount > -1 And X < Cx - 15 * (XRangeList - 1) Then
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
        ElseIf tmpGLShowCount < 0 Then 'List Head No Scroll Bar
            GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
        End If
    Else 'Thumb
        GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.PointerListHead = UIndexs + IndexList
    End If
    
    'Command16.Caption = TMPx2_PicThumb

    If ListPT.PTFull = 0 And GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab - 1 < Cy Then
'                    If tmpGLCount < CountGridList Then TMPx2_PicThumb = ListPT.PTX2

'MsgBox GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(rz + UIndexs).OpenThumb

        If GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(rz + UIndexs).OpenThumb = False Then
            TMPx2_PicThumb = 2
        Else
            TMPx2_PicThumb = ListPT.PTX2
        End If
        If ListPT.PTVis = 0 Then TMPx2_PicThumb = 2

        '''If (GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(rz + UIndexs).GLCount <=
        '''Int(CountList - Aab) \ GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(rz + UIndexs).GLRange) Then
        '''    TMPx2_PicThumb = ListPT.PTX2
        '''Else
        '''    TMPx2_PicThumb = 2
        '''    If (tmpGLShowCount = -1) Then TMPx2_PicThumb = ListPT.PTX2 'Else TMPx2_PicThumb = 2
        '''End If
        
        Picture1.Line (GridX(ShowFixCol(XPointerIndex)).GridLeft + TMPx2_PicThumb, _
        GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + 2)- _
        (Cx - 15 * (XRangeList - 1) - 1, _
        GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab - 1), vbWhite - 0, BF
        
        Picture1.Line (GridX(ShowFixCol(XPointerIndex)).GridLeft + TMPx2_PicThumb, _
        GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab - 17)- _
        (Cx - 15 * (XRangeList - 1) - 2, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab - 1), GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(rz + UIndexs).GLHeadColor, BF '&HFF8080, BF
    
        Picture1.FontBold = False
        TextEffect Picture1.hdc, GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSubType.ContGList(rz + UIndexs).GLCaption, _
        GridX(ShowFixCol(XPointerIndex)).GridLeft + 5 + TMPx2_PicThumb, _
        GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab - 14, _
        Cx - 15 * (XRangeList - 1) - 2 + 0, _
        GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * rz) + Aab, , GridXYData(ShowFixCol(XPointerIndex), YPointerIndex).GridSub.Alignment, 0
    End If
    
    'TextEffect ObjMe.hDC, GridXYData(TmpCutGridX, TmpCutGridY).GridSubType.ContGList(X + UIndexs).GLCaption & vbCrLf & "LLLLL", _
    GridX(TmpCutGridX).GridLeft + 5 + TMPx2_PicThumb, _
    ((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) - 14, _
    Cx - (15 * aXRangeList) - 2, _
    (((GridY(TmpCutGridY).GridTop + SellHeight_Def) + LoopLine) + 0) + 0, , GridXYData(TmpCutGridX, TmpCutGridY).GridSub.Alignment, 0
    
    'If GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList > GridY(YPointerIndex + 1).GridTop Then MsgBox "L"
    If GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList > Cy Then
        YyY = GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList - Cy
    Else
        YyY = 0
    End If
    'Command17.Visible = True
    'Command17.Caption = GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList - Cy
    Picture1.Line (GridX(ShowFixCol(XPointerIndex)).GridLeft + 2, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + 2)- _
    (Cx - 2, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList - YyY), vbWhite - 0, BF
    Test1 = ShowFixCol(XPointerIndex): Test2 = YPointerIndex
    DrawList Test1, Test2, Picture1, IndexList + 0, True
'    DrawList ShowFixCol(Test1), ShowFixCol(Test2), Picture1 ', IndexList + 0, True
    Picture1.Line (GridX(ShowFixCol(XPointerIndex)).GridLeft, Cy)-(Cx, Cy)


'Command16.Caption = Cy
Else
'RaiseEvent ClikSellSub("kk", 2, 2, 2)
End If

RaiseEvent ClikSellSub("kk", XPointerIndex, YPointerIndex, 2)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TmpCutGridX As Integer, TmpCutGridY As Integer
Dim nTmpCutGridX As Integer, nTmpCutGridY As Integer
Dim SGrid As Boolean, tmpXIndexGrid As Integer
Dim iXPointerIndex As Long, iYPointerIndex As Long
Dim Test1 As Integer, Test2 As Integer
Dim AdY As Single
Dim Ad As Integer, Fg As Integer
Dim YyY As Single

Command32.Caption = ""
'Exit Sub
If Text1.Visible = True Then Exit Sub

Command21.Caption = XPointerIndex

If XPointerIndex < 0 Then XPointerIndex = HeadSellOnFixCol(Abs(XPointerIndex) - 1)

If XPointerIndex = SetNewGrid.GridXCount Then XPointerIndex = SetNewGrid.GridXCount - 1
If YPointerIndex = SetNewGrid.GridYCount Then YPointerIndex = SetNewGrid.GridYCount - 1

Cx = GridX(FixdColnIndex(XPointerIndex)).GridLeft + GridX(FixdColnIndex(XPointerIndex)).GridWidth - 0
If YPointerIndex > SetNewGrid.GridYCount - 1 Then YPointerIndex = SetNewGrid.GridYCount - 1
Cy = GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight

'Scrol
If SubClikScrol = True Then
    MsgBox "This Line Error"
    'qwe = GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex
    '((((Y - TmpScrolY) - ((GridY(YPointerIndex).GridTop) + SellHeight_Def) - 13) / MovingY))
    GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex = _
    ((((Y - SubTmpScrolY) - ((GridY(YPointerIndex).GridTop) + SellHeight_Def) - 13) / SubMovingYs))
    
    'Command13.Caption = GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex
    
    If (Y - SubTmpScrolY) < GridY(YPointerIndex).GridTop + SellHeight_Def + 13 Then 'If < 0 From In List
        GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex = 0
        AdY = GridY(YPointerIndex).GridTop + SellHeight_Def + (13 * 1)
    ElseIf (Y - SubTmpScrolY) + 13 > Cy - 13 Then
        GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex = GridXYData(XPointerIndex, YPointerIndex).GridSubType.CountContGList - (GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLShowCount - 1) - 1
        AdY = Cy - (13 * 2)
    Else
        AdY = Y - SubTmpScrolY
    End If
'
    'Picture1.PaintPicture Image3.Picture, Cx - 15 * XRangeList, _
    AdY, , , 14 * 3, , 14, 13
    Static DFG As Integer
    If GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex > TMP_GLSubScrolIndex Or GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex < TMP_GLSubScrolIndex Then
        TMP_GLSubScrolIndex = GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex
        
        Picture1.Line (GridX(XPointerIndex).GridLeft + 2, (GridY(YPointerIndex).GridTop + SellHeight_Def) + 2)- _
        (Cx - 2, Cy - 1), vbWhite + 0, BF
        
        Picture1.Font.Bold = False
        Test1 = XPointerIndex: Test2 = YPointerIndex
        DrawList Test1, Test2, Picture1, , , , True
        
        'Picture1.CurrentY = 0
        'Picture1.Print DFG
        
        Picture1.Picture = Picture1.Image
        DFG = DFG + 1
        
    Else
    Picture1.Cls
    End If
    Command11.Caption = qwe & "---" & GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLSubScrolIndex
    Command10.Caption = DFG '"OPP " & GridXYData(XPointerIndex, YPointerIndex).GridSubType.CountContGList - (GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLShowCount - 1) - 1
    Picture1.PaintPicture Image3.Picture, Cx - 15, _
    AdY, , , 14 * 4, , 14, 13

    Exit Sub
End If

'Sub Scrol
If ClikScrol = True Then
    Aab = 20
    UIndexs = GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex

    tmpGLShowCount = GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.GLShowCount - 1
    'If tmpGLShowCount > -1 Then XRangeList = 2 Else XRangeList = 1
    
    'Clear
    If tmpGLShowCount > -1 Then XRangeList = 2 Else XRangeList = 1
    If GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList > Cy Then
        YyY = GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList - Cy
    Else
        YyY = 0
    End If
    
    Dim XxX As Single
    If GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.ContGList(IndexList).OpenThumb = False Then
        XxX = 0
    Else
        XxX = Get_Format(GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.GLFrmtStly, "?pt.x2|") - 1
    End If
    'sadsa dasd as Data dsa d
'    Command17.Caption = GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLFrmtStly

    Picture1.Line (GridX(FixdColnIndex(XPointerIndex)).GridLeft + 2 + XxX, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + 2)- _
    (Cx - 15 * XRangeList, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList - 1 - YyY), vbWhite, BF

    Command11.Enabled = True

    YListGrid = (Y - (SellHeight_Def + GridY(YPointerIndex).GridTop + Aab + 13 + (CountList * IndexList)))
    GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex = _
    (YListGrid - TmpScrolY) / MovingY
    
    Command13.Caption = GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex
    
    If GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex < 0 Then _
        GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex = 0
    
    If YListGrid - TmpScrolY < 0 Then
        Command14.Caption = "OK 1"
        AdY = GridY(YPointerIndex).GridTop + SellHeight_Def + Aab + 13 + CountList * IndexList
    ElseIf Int(YListGrid - TmpScrolY) + 13 > CountList - Aab + 1 - 13 * 2 Then
        AdY = GridY(YPointerIndex).GridTop + SellHeight_Def + CountList * (IndexList + 1) - 13 * 2 + 1 '+ (CountList - Aab + 1 - 13 * 3)
    Else
        Command14.Caption = "Cancel"
        AdY = Y - TmpScrolY
    End If

    Picture1.Font.Bold = False
    Test1 = FixdColnIndex(XPointerIndex): Test2 = YPointerIndex
    DrawList Test1, Test2, Picture1, 0 + IndexList, True, True
    
    If AdY + 13 > Cy Then YyY = Cy - (AdY + 0) Else YyY = 13
    If AdY < Cy Then
        Picture1.PaintPicture Image3.Picture, Cx - 15 * XRangeList, _
        AdY, , , 14 * 3, , 14, YyY
    End If
Exit Sub
End If


If Picture2.Visible = False And Picture3.Visible = False Then
    If ((X < FixedWidth And Y < FixedHeight) Or (X > FixedWidth And Y > FixedHeight)) And _
    Picture1.MousePointer = 2 Then Picture1.MousePointer = 0
        
    If Button = 1 Then
'        Command2.Caption = ""
        MDown = True
        If DragXY = False Then  'Drag
            iXPointerIndex = Hsc.Value: iYPointerIndex = Vsc.Value
            SearchPointer CBGrid, X, iXPointerIndex, Y, iYPointerIndex '-> Penggerak di jadikan satu
            
            If iXPointerIndex < 0 Then
                'iXPointerIndex = HeadSellOnFixCol(Abs(iXPointerIndex) - 1)
'                iXPointerIndex = GridX(HeadSellOnFixCol(Abs(iXPointerIndex) - 1)).GridRealOnPosisi
                iXPointerIndex = GridX(HeadSellOnFixCol(Abs(iXPointerIndex) - 1)).GridRealPosisi
            End If
'            If Px2 < 0 Then Px2 = HeadSellOnFixCol(Abs(Px2) - 1)
            
            TypeDragMove = True
            If TypeDrag = "XY" Then 'Drag In XY
                If iXPointerIndex = XPointerIndex And iYPointerIndex = YPointerIndex Then
                    Command29.Caption = "=="
                    Pxy.Px1 = XPointerIndex: Pxy.Py1 = YPointerIndex: Pxy.Px2 = XPointerIndex: Pxy.Py2 = YPointerIndex
                    DragXY = False
                ElseIf (iXPointerIndex < XPointerIndex And iYPointerIndex < YPointerIndex) Or (iXPointerIndex = XPointerIndex And iYPointerIndex < YPointerIndex) Then
                    Command29.Caption = "A <<"
                    Pxy.Px1 = iXPointerIndex: Pxy.Py1 = iYPointerIndex: Pxy.Px2 = XPointerIndex: Pxy.Py2 = YPointerIndex
                ElseIf (iXPointerIndex > XPointerIndex And iYPointerIndex < YPointerIndex) Or (iXPointerIndex > XPointerIndex And iYPointerIndex = YPointerIndex) Then
                    Command29.Caption = "B ><" & Now
                    Pxy.Px1 = XPointerIndex: Pxy.Py1 = iYPointerIndex: Pxy.Px2 = iXPointerIndex: Pxy.Py2 = YPointerIndex
                ElseIf (iXPointerIndex > XPointerIndex And iYPointerIndex > YPointerIndex) Or (iXPointerIndex = XPointerIndex And iYPointerIndex > YPointerIndex) Then
                    Command29.Caption = "C >>" & Now
                    Pxy.Px1 = XPointerIndex: Pxy.Py1 = YPointerIndex: Pxy.Px2 = iXPointerIndex: Pxy.Py2 = iYPointerIndex
                ElseIf (iXPointerIndex < XPointerIndex And iYPointerIndex > YPointerIndex) Or (iXPointerIndex < XPointerIndex And iYPointerIndex = YPointerIndex) Then
                    Command29.Caption = "D <>"
                    Pxy.Px1 = iXPointerIndex: Pxy.Py1 = YPointerIndex: Pxy.Px2 = XPointerIndex: Pxy.Py2 = iYPointerIndex
                End If
            ElseIf TypeDrag = "X" Then 'Drag Out  Y
                If (iXPointerIndex > XPointerIndex) Or (iXPointerIndex = XPointerIndex) Then
                    Pxy.Px1 = XPointerIndex: Pxy.Py1 = 0: Pxy.Px2 = iXPointerIndex: Pxy.Py2 = SellCountRow - 1
                Else
                    Pxy.Px1 = iXPointerIndex: Pxy.Py1 = 0: Pxy.Px2 = XPointerIndex: Pxy.Py2 = SellCountRow - 1
                End If
            ElseIf TypeDrag = "Y" Then 'Drag Out  X
                If (iYPointerIndex > YPointerIndex) Or (iYPointerIndex = YPointerIndex) Then
                    Pxy.Px1 = 0: Pxy.Py1 = YPointerIndex: Pxy.Px2 = SellCountColumn - 1: Pxy.Py2 = iYPointerIndex
                Else
                    Pxy.Px1 = 0: Pxy.Py1 = iYPointerIndex: Pxy.Px2 = SellCountColumn - 1: Pxy.Py2 = YPointerIndex
                End If
            End If
            CountTpXy_X = Pxy.Px2 - Pxy.Px1
            CountTpXy_Y = Pxy.Py2 - Pxy.Py1
    '        Exit Sub
        ElseIf DragXY = True Then  'Drag Jadi
            iXPointerIndex = Hsc.Value: iYPointerIndex = Vsc.Value
            SearchPointer CBGrid, X, iXPointerIndex, Y, iYPointerIndex '-> Penggerak di jadikan satu
            XPointerIndex = iXPointerIndex: YPointerIndex = iYPointerIndex
            
            If iXPointerIndex < 0 Then
            '    iXPointerIndex = HeadSellOnFixCol(Abs(iXPointerIndex) - 1)
'                iXPointerIndex = GridX(HeadSellOnFixCol(Abs(iXPointerIndex) - 1)).GridRealOnPosisi
                iXPointerIndex = GridX(HeadSellOnFixCol(Abs(iXPointerIndex) - 1)).GridRealPosisi
            End If
            
            TypeDragMove = True
                Pxy.Px1 = iXPointerIndex - TPxy.Px1: Pxy.Px2 = iXPointerIndex + TPxy.Px2
                    If Pxy.Px1 < 0 Then Pxy.Px1 = 0: Pxy.Px2 = (TPxy.Px2 + TPxy.Px1) '*
                    If Pxy.Px2 > SellCountColumn - 1 - HideCountFixCol Then Pxy.Px1 = (SellCountColumn - 1 - HideCountFixCol) - (TPxy.Px2 + TPxy.Px1): Pxy.Px2 = SellCountColumn - 1 - HideCountFixCol '*
                Pxy.Py1 = iYPointerIndex - TPxy.Py1: Pxy.Py2 = iYPointerIndex + TPxy.Py2
                    If Pxy.Py1 < 0 Then Pxy.Py1 = 0: Pxy.Py2 = (TPxy.Py2 + TPxy.Py1) '*
                    If Pxy.Py2 > SellCountRow - 1 Then Pxy.Py1 = (SellCountRow - 1) - (TPxy.Py2 + TPxy.Py1): Pxy.Py2 = SellCountRow - 1 '*
                    '* >> Penahanan kelebihan geser
    '        Exit Sub
        End If
        Command22.Caption = Pxy.Px1 & " " & Pxy.Px2
    
        If TypeDrag = "XY" Or TypeDrag = "X" Or TypeDrag = "[XY]" Then 'XY(X) or X Run to timer pergeseran
            If X > Picture1.ScaleWidth Then 'X+
                If SellLeft(ShowFixCol(GridRightX)) + SellWidth(ShowFixCol(GridRightX)) > Picture1.ScaleWidth Then _
                   HTPxy.Px2 = 1
            ElseIf X < FixedWidth Then   'X-
                If Hsc.Value <> 0 Then _
                   HTPxy.Px1 = 1
            Else                            'X
                HTPxy.Px1 = 0
                HTPxy.Px2 = 0
            End If
        End If
        
'        Command21.Caption = HTPxy.Px2
        
        If TypeDrag = "XY" Or TypeDrag = "Y" Or TypeDrag = "[XY]" Then 'XY(Y) or Y Run to timer pergeseran
            If Y > Picture1.ScaleHeight Then 'Y+
                If SellTop(GridDownY) + SellHeight(GridDownY) > Picture1.ScaleHeight Then _
                   HTPxy.Py2 = 1
            ElseIf Y < FixedHeight Then    'Y-
                If Vsc.Value <> 0 Then _
                   HTPxy.Py1 = 1
            Else                             'Y
                HTPxy.Py1 = 0
                HTPxy.Py2 = 0
            End If
        End If
        If (X > 0 And X < Picture1.ScaleWidth) And (Y > 0 And Y < Picture1.ScaleHeight) Then
            HTPxy.Px1 = 0
            HTPxy.Px2 = 0
            HTPxy.Py1 = 0
            HTPxy.Py2 = 0
    
            TypeDragMove = False
        End If
        Timer1.Enabled = TypeDragMove
        
        If HTPxy.Px1 = 0 And HTPxy.Px2 = 0 And HTPxy.Py1 = 0 And HTPxy.Py2 = 0 Then
''            Picture1.Cls >>>>>>>>> Sementara Dihilangkan Untuk Program Billing
            If Pusing = True Then
                GridXY Picture1, Hsc.Value, Vsc.Value
                Picture1.Picture = Picture1.Image
                Pusing = False
            End If
            
            'Command7.Caption = Pusing
            'Picture1.Cls
            DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On1            DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
                If iXPointerIndex <> XPointerIndex Or iYPointerIndex <> YPointerIndex Then _
                    DrawInvert Picture1, SellLeft(FixdColnIndex(XPointerIndex)) + 3, SellTop(YPointerIndex) + 3, SellLeft(FixdColnIndex(XPointerIndex)) + SellWidth(FixdColnIndex(XPointerIndex)) - 3, SellTop(YPointerIndex) + SellHeight(YPointerIndex) - 3
            DrawInvert Picture1, SellLeft(FixdColnIndex(XPointerIndex)) + 3, 3, SellLeft(FixdColnIndex(XPointerIndex)) + SellWidth(FixdColnIndex(XPointerIndex)) - 3, FixedHeight - 3
            DrawInvert Picture1, 2, SellTop(YPointerIndex) + 2, FixedWidth - 2, SellTop(YPointerIndex) + SellHeight(YPointerIndex) - 2
        End If
        'kkkkkkk k   kkkkkkkkkkk
        Command2.Caption = TypeDragMove
    End If
End If

'Events Grid X, Y Range---------------------------------------------
If MDown = False And ((X > FixedWidth And Y < FixedHeight) Or Fx = True) And Fy = False Then
    If Picture2.Visible = False And Picture3.Visible = False And Text1.Visible = False Then
        If X > FixedWidth And Y < FixedHeight Then _
        MouseX CBGrid, Picture1, Picture3, X, Y, Hsc.Value   'Else Picture1.MousePointer = 0
    Else
        If Button = 1 Then
            Picture3.left = X
            XMovGrid = (SellLeft(FixdColnIndex(XIndexGrid)) + SellWidth(FixdColnIndex(XIndexGrid))) - X
            Fx = True: Fy = False
        End If
    End If
ElseIf MDown = False And ((X < FixedWidth And Y > FixedHeight) Or Fy = True) And Fx = False Then
    If Picture2.Visible = False And Picture3.Visible = False And Text1.Visible = False Then
        If X < FixedWidth And Y > FixedHeight Then _
        MouseY CBGrid, Picture1, Picture2, X, Y, Vsc.Value 'Else Picture1.MousePointer = 0
    Else
        If Button = 1 Then
            Picture2.top = Y
            YMovGrid = ((SellTop(YIndexGrid) + Abs(SellHeight(YIndexGrid))) - Y)
            Fx = False: Fy = True
        End If
    End If
End If

RaiseEvent MouseMoveSell(iXPointerIndex, iYPointerIndex, Button, Shift, X, Y)

Command9.Caption = X & ":" & Y ' - 210
'----------------------------------------------------------------------
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Test1 As Integer, Test2 As Integer

If ClikScrol = True Or SubClikScrol = True Then
    UIndexs = GridXYData(FixdColnIndex(XPointerIndex), YPointerIndex).GridSubType.GLSubScrolIndex
 
    Test1 = FixdColnIndex(XPointerIndex): Test2 = YPointerIndex
    DrawList Test1, Test2, Picture1, IndexList, True
    Picture1.Picture = Picture1.Image
    List1.AddItem MovingY

End If

SubClikScrol = False
ClikScrol = False
Timer1.Enabled = False
MDown = False
TypeDragMove = False
If Picture2.Visible = True Or Picture3.Visible = True Then 'Range
    If Picture3.Visible = True Then MouseUpX CBGrid, Picture1, Picture3
    If Picture2.Visible = True Then MouseUpY CBGrid, Picture1, Picture2
    
    'Picture1.Picture = Nothing
    'Picture1.Cls
    GridXY Picture1, Hsc.Value, Vsc.Value
    Picture1.Picture = Picture1.Image
    'If GridType <> 1 And GridType <> 3 Then _

        DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On11        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    
    Fx = False
    Fy = False
Else
    If SizeY(CBGrid, X, Y) = True Then 'Open Sub Grid
        'Picture1.Picture = Nothing
        'Picture1.Cls
        GridXY Picture1, Hsc.Value, Vsc.Value
        Picture1.Picture = Picture1.Image
        If CheckType(GridType, 0) = False Then _
            DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On12            DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    Else
        If DragXY = True Then 'Clear Tmp Drag
'            Picture1.Cls
            DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On13            DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
                DrawInvert Picture1, SellLeft(FixdColnIndex(XPointerIndex)) + 3, SellTop(YPointerIndex) + 3, SellLeft(FixdColnIndex(XPointerIndex)) + SellWidth(FixdColnIndex(XPointerIndex)) - 3, SellTop(YPointerIndex) + SellHeight(YPointerIndex) - 3
        End If
    End If
    RaiseEvent MouseUpSell(XPointerIndex, YPointerIndex, Button, Shift, X, Y)
End If
End Sub

Private Sub Picture1_Paint()
'        DrawInvertToGrid cbgrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
'DoEvents

End Sub

Private Sub Timer1_Timer()
Dim TmpPxy As PointInvert
Static DelayInvert As Integer, Hi As Boolean

'Exit Sub
DelayInvert = DelayInvert + 1
DelayOver = 1
Timer1.Interval = 100

'Update
Command3.Caption = Second(Time) 'DelayInvert
    Command7.Caption = TPxy.Px1
'Command5.Caption = ""

TmpPxy = Pxy
If DelayInvert >= DelayOver Then
    If HTPxy.Px2 = 1 Then 'X2
        If DragXY = False Then
            If Pxy.Px2 > XPointerIndex Then
                Pxy.Px2 = GridRightX
            Else
                Pxy.Px1 = GridRightX 'update err
                    If Pxy.Px1 > XPointerIndex Then
                        Pxy.Px1 = XPointerIndex
                        Pxy.Px2 = GridRightX
                    End If
            End If
            CountTpXy_X = Pxy.Px2 - Pxy.Px1
        Else
            If GridRightX > TPxy.Px1 Then
                If Pxy.Px2 <> SellCountColumn - 1 - HideCountFixCol Then
                    'Pxy.Px1 = GridRightX - TPxy.Px1:
                    Pxy.Px2 = GridRightX + TPxy.Px2
                        If Pxy.Px2 >= SellCountColumn - 1 - HideCountFixCol Then Pxy.Px2 = SellCountColumn - 1 - HideCountFixCol
                    Pxy.Px1 = Pxy.Px2 - (CountTpXy_X)
                    'List2.AddItem Pxy.Px1
                    If Pxy.Px1 < 0 Then MsgBox ""
                ''Else
                'here here here
                    'List2.AddItem Pxy.Px1 & " " & Pxy.Px2 & " - " & TPxy.Px2
                End If
            
            End If
        End If
        
        If SellLeft(FixdColnIndex(GridRightX)) + SellWidth(FixdColnIndex(GridRightX)) >= Picture1.ScaleWidth And Hsc.Value <> Hsc.Max Then _
        Hsc.Value = Hsc.Value + 1 Else HTPxy.Px2 = 0 'masih ada err
    ElseIf HTPxy.Px1 = 1 Then 'X1
        Hsc.Value = Hsc.Value - 1 'untuk stop di label Stop_Hsc.Value
        If DragXY = False Then
            If Pxy.Px1 < XPointerIndex Then
                Pxy.Px1 = Hsc.Value
            Else
                Pxy.Px2 = Hsc.Value
                    If Pxy.Px2 < XPointerIndex Then
                        Pxy.Px1 = Hsc.Value
                        Pxy.Px2 = XPointerIndex
                    End If
            End If
        Else
            If Pxy.Px1 <> 0 Then _
            Pxy.Px1 = Hsc.Value - TPxy.Px1: Pxy.Px2 = Hsc.Value + TPxy.Px2
        End If
    ElseIf HTPxy.Px1 = 0 And HTPxy.Px2 = 0 Then
        'If Picture1.Picture = 0 Then
        '    Picture1.Cls
        '    GridXY Picture1, Hsc.Value, Vsc.Value
        '    Picture1.Picture = Picture1.Image
        'End If
    End If
    
    If HTPxy.Py2 = 1 Then
        If DragXY = False Then
            If Pxy.Py2 > YPointerIndex Then
                Pxy.Py2 = GridDownY
            Else
                Pxy.Py1 = GridDownY 'update
                    If Pxy.Py1 > YPointerIndex Then
                        Pxy.Py1 = YPointerIndex
                        Pxy.Py2 = GridDownY
                    End If
            End If
            CountTpXy_Y = Pxy.Py2 - Pxy.Py1
        Else
            If GridDownY > TPxy.Py1 Then
                If Pxy.Py2 <> SellCountRow - 1 Then '_
'                Pxy.Py1 = GridDownY - TPxy.Py1: Pxy.Py2 = GridDownY + TPxy.Py2
                    Pxy.Py2 = GridDownY + TPxy.Py2
                        If Pxy.Py2 >= SellCountRow - 1 Then Pxy.Py2 = GridDownY
                    Pxy.Py1 = Pxy.Py2 - (CountTpXy_Y)
                End If
            End If
        End If
        
        If SellTop(GridDownY) + SellHeight(GridDownY) >= Picture1.ScaleHeight And Vsc.Value <> Vsc.Max Then _
            Vsc.Value = Vsc.Value + 1 Else HTPxy.Py2 = 0
    ElseIf HTPxy.Py1 = 1 Then
        Vsc.Value = Vsc.Value - 1 'untuk stop di label Stop_Vsc.Value
            
        If DragXY = False Then
            If Pxy.Py1 < YPointerIndex Then
                Pxy.Py1 = Vsc.Value
            Else
                Pxy.Py2 = Vsc.Value
                    If Pxy.Py2 < YPointerIndex Then
                        Pxy.Py1 = Vsc.Value
                        Pxy.Py2 = YPointerIndex
                    End If
            End If
        Else
            If Pxy.Py1 <> 0 Then _
            Pxy.Py1 = Vsc.Value - TPxy.Py1: Pxy.Py2 = Vsc.Value + TPxy.Py2
        End If
    
    ElseIf HTPxy.Py1 = 0 And HTPxy.Py2 = 0 Then
    End If
















'Exit Sub




'Text4 = Text4 & HTPxy.Px1 & " "
    If HTPxy.Px1 = 0 And HTPxy.Px2 = 0 And HTPxy.Py1 = 0 And HTPxy.Py2 = 0 Then
        'If Picture1.Picture = 0 Then
        '    Picture1.Cls
        '    GridXY Picture1, Hsc.Value, Vsc.Value
        '    Picture1.Picture = Picture1.Image
        'End If
    End If
    If HTPxy.Px1 = 1 Or HTPxy.Px2 = 1 Or HTPxy.Py1 = 1 Or HTPxy.Py2 = 1 Then
        If Hsc.Value = 0 Then HTPxy.Px1 = 0  'Label : Stop_Hsc.Value
        If Vsc.Value = 0 Then HTPxy.Py1 = 0  'Label : Stop_Hsc.Value



'            If Pxy.Px2 = SellCountColumn - 1 Then HTPxy.Px2 = 0
            'If Pxy.Py2 = SellCountRow - 1 Then HTPxy.Py2 = 0
'Command5.Caption = Pxy.Py2 & "  " & GridDownY
            If Picture1.Picture <> 0 Then
            'Picture1.Picture = Nothing
            'MsgBox """"
            End If
            Picture1.Cls
            GridXY Picture1, Hsc.Value, Vsc.Value
            'If Pxy.Px1 < 0 Then MsgBox ""
            DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On2            DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
            'Picture1.Picture = Picture1.Image
            
            Command5.Caption = Second(Time)
            Command6.Caption = 1
                        Pusing = True

    'Me.Caption = Now
    Else
            Command6.Caption = 0
            ''If Pusing = True Then
                Picture1.Cls
                GridXY Picture1, Hsc.Value, Vsc.Value
                Picture1.Picture = Picture1.Image
                DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On3                DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2

                Pusing = False
    End If
        List2.AddItem TPxy.Px1
                            
        If GridX(FixdColnIndex(Pxy.Px1 + TPxy.Px1)).GridLeft + GridX(FixdColnIndex(Pxy.Px1 + TPxy.Px1)).GridWidth > Picture1.ScaleWidth Then
            TypeDragMove = False ' nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn
            Hsc.Value = SearchOverGrid(Pxy.Px1 + TPxy.Px1, False)  'XYZ + 1
                If GridX(ShowFixCol(Pxy.Px1 + TPxy.Px1)).GridLeft + GridX(ShowFixCol(Pxy.Px1 + TPxy.Px1)).GridWidth > Picture1.ScaleWidth Then Hsc.Value = Hsc.Value + 1 'pertimbangan
            TypeDragMove = True
        End If
        If GridY(Pxy.Py1 + TPxy.Py1).GridTop + GridY(Pxy.Py1 + TPxy.Py1).GridHeight > Picture1.ScaleHeight Then
            TypeDragMove = False ' nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn
            Vsc.Value = SearchOverGrid(Pxy.Py1 + TPxy.Py1, True)  'XYZ + 1
                If GridY(Pxy.Py1 + TPxy.Py1).GridTop + GridY(Pxy.Py1 + TPxy.Py1).GridHeight > Picture1.ScaleHeight Then Vsc.Value = Vsc.Value + 1 'pertimbangan
            TypeDragMove = True
        End If

DelayInvert = 0
End If
        
        Command1.Caption = Pxy.Px1 & " bbbb"
End Sub

Private Sub Picture6_Click()
Picture6.Visible = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Cancels As Boolean, Texts As String
Dim NoHideText As Boolean

Texts = Text1.Text
RaiseEvent KeyDownInput(XPointerIndex, YPointerIndex, KeyCode, Shift, Texts, NoHideText)
If NoHideText = True Then
'    Picture1.Picture = tmpPicture1.Picture
'    Picture1.Picture = Picture1.Image
    Exit Sub
End If
Text1.Text = Texts

If KeyCode = 13 Then
    If (DblX > FixedWidth And DblY < FixedHeight) Then
        FixedColsText(ShowFixCol(XPointerIndex)) = Text1.Text
        Text1.Visible = False
    ElseIf (DblX < FixedWidth And DblY > FixedHeight) Then
        FixedRowsText(YPointerIndex) = Text1.Text
        Text1.Visible = False
    Else
        Texts = Text1.Text
        RaiseEvent KeyEnter(XPointerIndex, YPointerIndex, Texts, Cancels)
        If Cancels = False Then
            SellText(FixdColnIndex(XPointerIndex), YPointerIndex) = Texts
        End If
        Text1.Visible = False
    End If
    
    Picture1.SetFocus
    Fx = False
    Fy = False
    
    ''GridXY Picture1, Hsc.Value, Vsc.Value >MyBe To Del
    ''Picture1.Picture = Picture1.Image
'    Picture1.Picture = tmpPicture1.Picture
    'If GridType <> 1 And GridType <> 3 Then
    ''DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2 >MyBe To Del
    
    NoHideText = False
End If

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.MousePointer = 0
Fx = False
Fy = False
End Sub

Private Sub Timer2_Timer()

If ShowOnConHeadSell > 0 Then
Command28.Caption = ""
For X = 0 To ShowOnConHeadSell - 1
    Command28.Caption = Command28.Caption & HeadSellOnFixCol(X) & "." & GridX(HeadSellOnFixCol(X)).GridLeft & " "
Next X
End If
Command2.Caption = ClickScroll 'Pxy.Px1 & " " & TPxy.Px1
Command27.Caption = Pxy.Px1 & " " & Pxy.Px2 & " Count: " & Pxy.Px2 - Pxy.Px1 & " " & CountTpXy_X
End Sub


Sub DrawPicGrid(ObjMe As Object, IndexXY As Integer, TmpCutGridX As Integer, TmpCutGridY As Integer)
Dim IndexX As Integer, IndexY As Integer
Dim ax As Integer

If IndexXY < 0 Then Exit Sub

If SetNewGrid.GridSizePic.SellIconX1 < 1 Then SetNewGrid.GridSizePic.SellIconX1 = 3
If SetNewGrid.GridSizePic.SellIconX2 < 1 Then SetNewGrid.GridSizePic.SellIconX2 = 20
If SetNewGrid.GridSizePic.SellIconY2 < 1 Then SetNewGrid.GridSizePic.SellIconY2 = 20
If SetNewGrid.GridSizePic.SellIconY1 < 1 Then SetNewGrid.GridSizePic.SellIconY1 = ((SetNewGrid.GridSize.SellHeight_Def - SetNewGrid.GridSizePic.SellIconY2) / 2) + 1

If PicErr = False Then
    IndexY = IndexXY \ (SetNewGrid.GridSizePic.SellIconPicContColms + 1)
    IndexX = IndexXY - (IndexY * (SetNewGrid.GridSizePic.SellIconPicContColms + 1))
    
    ax = SetNewGrid.GridSizePic.SellIconX2
    If SetNewGrid.GridSizePic.SellIconX2 + SetNewGrid.GridSizePic.SellIconX1 > GridX(TmpCutGridX).GridWidth Then _
        ax = GridX(TmpCutGridX).GridWidth - SetNewGrid.GridSizePic.SellIconX1 'MsgBox "O"
    If ax < 0 Then Exit Sub
    'If tmpPictureGrid.Handle = Nothing Then MsgBox LLL
'    MsgBox tmpPictureGrid.Height
    On Error GoTo 10
    'MsgBox tmpPictureGrid
    ObjMe.PaintPicture tmpPictureGrid, GridX(TmpCutGridX).GridLeft + SetNewGrid.GridSizePic.SellIconX1, _
    GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePic.SellIconY1, _
    , , ax * IndexX, _
    SetNewGrid.GridSizePic.SellIconY2 * IndexY, ax - 0, SetNewGrid.GridSizePic.SellIconY2 - 0
Else
    ObjMe.PaintPicture tmpPictureGrid, GridX(TmpCutGridX).GridLeft + SetNewGrid.GridSizePic.SellIconX1, _
    GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePic.SellIconY1 ', _
    , , SetNewGrid.GridSizePic.SellIconX2 * IndexX, _
    SetNewGrid.GridSizePic.SellIconY2 * IndexY, SetNewGrid.GridSizePic.SellIconX2 - 0, SetNewGrid.GridSizePic.SellIconY2 - 0
End If

Exit Sub
10:
'MsgBox "Not SLinkPictureXY to PictureXY"
'jkj j jkjk jk kj
    Set tmpPictureGrid = Image9.Picture
    ObjMe.PaintPicture tmpPictureGrid, GridX(TmpCutGridX).GridLeft + SetNewGrid.GridSizePic.SellIconX1, _
    GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePic.SellIconY1, _
    , , ax * IndexX, _
    SetNewGrid.GridSizePic.SellIconY2 * IndexY, ax - 0, SetNewGrid.GridSizePic.SellIconY2 - 0
End Sub

Sub PictureIndex(ObjMe As Object, Colms As Integer, tX1 As Integer, tY1 As Integer, tX2 As Integer, tY2 As Integer)
Dim IndexX As Integer, IndexY As Integer

    IndexY = IndexXY \ (Colms + 1)
    IndexX = IndexXY - (IndexY * (Colms + 1))


    Dim Ax2 As Integer
    
    Ax2 = tX2
    If tX2 + tX1 > GridX(TmpCutGridX).GridWidth Then _
        Ax2 = GridX(TmpCutGridX).GridWidth - tX1
    
    ObjMe.PaintPicture tmpPictureGrid, GridX(TmpCutGridX).GridLeft + tX1, _
    GridY(TmpCutGridY).GridTop + tY1, _
    , , Ax2 * IndexX, _
    tY2 * IndexY, Ax2 - 0, tY2 - 0
End Sub
















'Sub iDrawXY(ObjMe As Object, IndexX As Integer, IndexY As Integer, TmpCutGridX As Integer, TmpCutGridY As Integer)
'If SetNewGrid.GridSizePicSub.SellIconX1 < 1 Then SetNewGrid.GridSizePicSub.SellIconX1 = 3
'If SetNewGrid.GridSizePicSub.SellIconY1 < 1 Then SetNewGrid.GridSizePicSub.SellIconY1 = (SetNewGrid.GridSize.SellHeight_Def - 15) / 2
'If SetNewGrid.GridSizePicSub.SellIconX2 < 1 Then SetNewGrid.GridSizePicSub.SellIconX2 = 20
'If SetNewGrid.GridSizePicSub.SellIconY2 < 1 Then SetNewGrid.GridSizePicSub.SellIconY2 = 20
'
'On Error GoTo 10
'    ObjMe.PaintPicture PictureXY, GridX(TmpCutGridX).GridLeft + SetNewGrid.GridSizePicSub.SellIconX1, _
'    GridY(TmpCutGridY).GridTop + SetNewGrid.GridSizePicSub.SellIconY1, _
'    , , (SetNewGrid.GridSizePicSub.SellIconX2 * IndexX - IndexX) + 0, _
'    (SetNewGrid.GridSizePicSub.SellIconY2 * IndexY - IndexY) + 1, SetNewGrid.GridSizePicSub.SellIconX2 - 0, SetNewGrid.GridSizePicSub.SellIconY2 - 0
'Exit Sub
'10
'MsgBox "Not SLinkPictureXY to PictureXY"
'End
'End Sub




Private Sub Timer3_Timer()
'****
GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(IndexGL).GLScrolIndex = GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(IndexGL).GLScrolIndex + 1
GridXY Picture1, Hsc.Value, Vsc.Value
'Picture1.Picture = Picture1.Image
Timer3.Interval = 1
End Sub

Private Sub UserControl_GotFocus()
If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub UserControl_Initialize()
Dim nLen As Long

'lblToolTipText.Caption = "F1 Show Text " '& vbCrLf & "F2 Show Text"
'lblToolTipText.Caption = "F1 " & vbCrLf & "F2"
'nLen = Len(lblToolTipText.Caption)

'MsgBox picToolTipText.TextHeight(lblToolTipText.Caption)
'picToolTipText.Width

Form_Mini_Normal picScrollMove

Image3.Picture = LoadPicture(App.Path & "\A.BMP")

Picture1.left = 3
Picture1.top = 3
Picture1.Width = UserControl.ScaleWidth - 3
Picture1.Height = UserControl.ScaleHeight - 3

''XXpvDrawGrid Picture5.hDC, 0, 0, 100, 100, vbBlue, vbRed, 0
RangeEndHead = 200
RangeEndHead_Pos = RangeEndHead
RangeEndHead_PosAuto = 3
HeadCountAuto = False


Add 1, 1 '99, 99

SetNewGrid.GridLenkapX.GridBackColGra(0) = Picture1.BackColor
SetNewGrid.GridLenkapX.GridForeColor = Picture1.ForeColor
SetNewGrid.GridLenkapY.GridBackColGra(0) = Picture1.BackColor
SetNewGrid.GridLenkapY.GridForeColor = Picture1.ForeColor
SetNewGrid.GridXYBackColor = Picture1.BackColor
SetNewGrid.GridXYBackColorSub = vbWhite
SetNewGrid.GridXYForeColor = Picture1.ForeColor
SetNewGrid.GridXYForeColorSub = Picture1.ForeColor








'SellIconFilePicture = "D:\Pic.bmp"
'PictureGrid = LoadPicture(SetNewGrid.GridFilePicture)

PictureSubGridY = Image2.Picture

'Image4.Picture = LoadPicture(SetNewGrid.GridFilePicture)
'Image6.Picture = LoadPicture(SetNewGrid.GridFilePicture)
'SLinkPictureSub Image2
'SLinkPictureXY Image4

SetNewGrid.GridSizePic.SellIconPicContColms = 6
SetNewGrid.GridSizePic.SellIconPicContRows = 6
'SetNewGrid.GridSizePic.SellIconX1 = 50
SetNewGrid.GridSizePic.SellIconX2 = 20
'SetNewGrid.GridSizePic.SellIconY2 = 30

'GridX(5).GPicturePut = True
'GridX(0).GPicturePut = True
''GridXYData(0, 3).GridXYPicIndex = 3
''GridXYData(0, 1).GridXYPicIndex = 8
'GridXYData(0, 1).GridSubType = 1
'GridX(7).GPicturePut = True

'MsgBox GridXYData(5, 0).GridXYPicIndex









'FixedWidth = 25
'FixedHeight = 20
SetNewGrid.GridSize.GDRangeX = 25
SetNewGrid.GridSize.GDRangeY = 20

SellWidth_Def = 80
SellHeight_Def = 25

TableWidth = Picture1.ScaleWidth
TableHeight = Picture1.ScaleHeight

RangePicSubX1 = 3
RangePicSubY1 = 3
RangePicSubX2 = 9
RangePicSubY2 = 9

'GridXYData(1, 150).Grid.BackColor = vbBlue
'GridXYData(1, 150).Grid.GColorDefault(0) = True

'SellBackColor(1, 0) = 0
'GridXYData(1, 0).Grid.BackColor = True = 0
'GridXYData(1, 0).Grid.GColorDefault(0) = True
'GridXYData(1, 0).Grid.BackColor = 0
'GridXYData(2, 0).Grid.GColorDefault(0) = True
'GridXYData(2, 0).Grid.BackColor = 0


GridXY Picture1, Hsc.Value, Vsc.Value
Picture1.Picture = Picture1.Image

'Picture1.Picture = PictureXY.Picture
'DrawingGrid Picture1, Hsc.Value, Vsc.Value

'Vsc.Value = 150
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'BackColor = PropBag.ReadProperty("BackColor", &H8000000F)

'FixedWidth = PropBag.ReadProperty("FixedWidth", 25)
'FixedHeight = PropBag.ReadProperty("FixedHeight", 20)

'GridXY Picture1, Hsc.Value, Vsc.Value
'Picture1.Picture = Picture1.Image

End Sub

Private Sub UserControl_Resize()
Picture1.left = 1
Picture1.top = 1
Picture1.Width = (UserControl.ScaleWidth - Vsc.Width) - 3
    Hsc.Width = Picture1.Width
    Vsc.left = Picture1.left + Picture1.Width + 1
    Vsc.top = Picture1.top
Picture1.Height = (UserControl.ScaleHeight - Hsc.Height) - 3
    Vsc.Height = Picture1.Height
    
    Hsc.left = Picture1.left
    Hsc.top = Picture1.top + Picture1.Height + 1
TableWidth = Picture1.ScaleWidth
TableHeight = Picture1.ScaleHeight

'DoEvents
GridXY Picture1, Hsc.Value, Vsc.Value
Picture1.Picture = Picture1.Image

GridFullWidth = Picture1.Width
GridFullHeight = Picture1.Height

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'PropBag.WriteProperty "BackColor", BackColor, &H8000000F

PropBag.WriteProperty "FixedWidth", FixedWidth, 25
PropBag.WriteProperty "FixedHeight", FixedHeight, 20



'GridXY Picture1, Hsc.Value, Vsc.Value
'Picture1.Picture = Picture1.Image
'^&*(
End Sub

Private Sub HscHead_Change()
Dim TmpNoDrawing As Boolean

GridLeftX = Hsc.Value
If TypeDragMove = False Then
    If Picture1.Picture <> 0 Then Picture1.Picture = Nothing
    RaiseEvent ChangeHBefore
    
Picture1.Cls  'Picture1.BackColor = UserControl.BackColor
GridXY Picture1, Hsc.Value, Vsc.Value, , , , TmpNoDrawing ', , 1
    
    If CheckType(GridType, 0) = False Then _
        DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On15        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    
    RaiseEvent ChangeHAfter
End If
ClickScroll = True
End Sub

Private Sub HscHead_Scroll()
DoEvents
GridLeftX = Hsc.Value
If Picture1.Picture <> 0 Then Picture1.Picture = Nothing

Picture1.Cls  'Picture1.BackColor = UserControl.BackColor
GridXY Picture1, Hsc.Value, Vsc.Value
    If Timer1.Enabled = False Then _
    If CheckType(GridType, 0) = False Then _
        DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On On        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
RaiseEvent ScrollH
ClickScroll = True
'Picture1.SetFocus
End Sub

Private Sub Hsc_Change()
Dim TmpNoDrawing As Boolean

GridLeftX = Hsc.Value
If TypeDragMove = False Then
    If Picture1.Picture <> 0 Then Picture1.Picture = Nothing
    RaiseEvent ChangeHBefore
    
Picture1.Cls  'Picture1.BackColor = UserControl.BackColor
GridXY Picture1, Hsc.Value, Vsc.Value, , , , TmpNoDrawing ', , 1
    
    If CheckType(GridType, 0) = False Then _
        DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On15        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    
    RaiseEvent ChangeHAfter
End If
ClickScroll = True
'Picture1.SetFocus
End Sub

Private Sub Hsc_Scroll()
DoEvents
GridLeftX = Hsc.Value
If Picture1.Picture <> 0 Then Picture1.Picture = Nothing


Picture1.Cls  'Picture1.BackColor = UserControl.BackColor
GridXY Picture1, Hsc.Value, Vsc.Value
    If Timer1.Enabled = False Then _
    If CheckType(GridType, 0) = False Then _
        DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On On        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
RaiseEvent ScrollH
ClickScroll = True
'Picture1.SetFocus
End Sub

Private Sub Hsc_LostFocus()
    'HeadOnHSC
End Sub

Private Sub Vsc_Change()
'Exit Sub
GridUpY = Vsc.Value
If TypeDragMove = False Then
    If Picture1.Picture <> 0 Then Picture1.Picture = Nothing
    RaiseEvent ChangeVBefore

Picture1.Cls  'Picture1.BackColor = UserControl.BackColor
GridXY Picture1, Hsc.Value, Vsc.Value
    If CheckType(GridType, 0) = False Then _
        DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On16        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
    
    RaiseEvent ChangeVAfter
End If
ClickScroll = True

'Picture1.SetFocus
End Sub

Private Sub Vsc_Scroll()
'Picture1.Enabled = True
DoEvents
'GridUpY = Vsc.Value
GridUp = Vsc.Value
If Picture1.Picture <> 0 Then Picture1.Picture = Nothing

Picture1.Cls
GridXY Picture1, Hsc.Value, Vsc.Value
    If CheckType(GridType, 0) = False Then _
        DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
''On On        DrawInvertToGrid CBGrid, Picture1, Pxy, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
RaiseEvent ScrollV
ClickScroll = True
'Picture1.SetFocus
End Sub

'Sub FormSetData(CFileAll As Integer)
'    SetData CFileAll
'End Sub
'Sub FormGetData(CFileAll As Integer)
'    GetData CFileAll
'End Sub
'Sub RefreshPic()
'    GridXY Picture1, 0, 0   'Hsc.Value, Vsc.Value
'    MsgBox Me.Tag
'End Sub


Sub Pic6Show(Txt As String)
    Picture6.Visible = True
    Picture6.Cls
    TextEffect Picture6.hdc, Txt, 0, 0, 0, 0, 0
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------------------------------















































Sub XXXXMouseX(tBGrid As Integer, ObjMe As Object, ObjMe2 As Object, X As Single, Y As Single, MdlCountX As Integer)
Dim TmpCutGridX As Integer

'If Y < FixedHeight Then
    TmpCutGridX = MdlCountX
    Do
        If TmpCutGridX < SellCountColumn Then
            Cx = SellLeft(TmpCutGridX) + SellWidth(TmpCutGridX)
            If X < Cx Or TmpCutGridX > SellCountColumn - 1 Then
                If X > Cx - 5 Or X < SellLeft(TmpCutGridX) + 5 Then
                    If X < SellLeft(TmpCutGridX) + 5 Then az = 1
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
Command21.Caption = XIndexGrid
'MsgBox ShowFixCol
End Sub


Sub TMPMouseX(tBGrid As Integer, ObjMe As Object, ObjMe2 As Object, X As Single, Y As Single, MdlCountX As Integer)
Dim TmpCutGridX As Integer

'If Y < FixedHeight Then
    TmpCutGridX = MdlCountX
    Do
        If TmpCutGridX < SellCountColumn - HideCountFixCol Then
            Cx = SellLeft(FixdColnIndex(TmpCutGridX)) + SellWidth(FixdColnIndex(TmpCutGridX))
            If X < Cx Or TmpCutGridX > SellCountColumn - 1 Then
                If X > Cx - 5 Or X < SellLeft(ShowFixCol(TmpCutGridX)) + 5 Then
                    If X < SellLeft(FixdColnIndex(TmpCutGridX)) + 5 Then az = 1
                        If TmpCutGridX - az < 0 Then az = 0
                        
                        ObjMe.MousePointer = 2
                        XIndexGrid = ShowFixCol(TmpCutGridX - az)
                        XPointerIndex = ShowFixCol(TmpCutGridX - az)
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
Command21.Caption = XIndexGrid
'MsgBox ShowFixCol(TmpCutGridX)
End Sub

'Private Sub Picture Events -------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub Original_MouseX(tBGrid As Integer, ObjMe As Object, ObjMe2 As Object, X As Single, Y As Single, MdlCountX As Integer)
Dim TMPTmpCutGridX As Integer, TmpCutGridX As Integer
Dim MyBlock As Boolean, UpBlock As Boolean

'If Y < FixedHeight Then
    TmpCutGridX = MdlCountX
'    TMPTmpCutGridX = HscHead.Value
    Do
        If TmpCutGridX < SellCountColumn - HideCountFixCol Then
            'Bisa juga di gunakan dari jumlah lebar pada headnya
            If UpBlock = False And ShowOnConHeadSell > 0 Then
                MyBlock = True
                Cx = SellLeft(HeadSellOnFixCol(TMPTmpCutGridX)) + SellWidth(HeadSellOnFixCol(TMPTmpCutGridX))
                If X < Cx Or TmpCutGridX > SellCountColumn - 1 Then
                    If X > Cx - 5 Or X < SellLeft(HeadSellOnFixCol(TMPTmpCutGridX)) + 5 Then
                        If X < SellLeft(HeadSellOnFixCol(TMPTmpCutGridX)) + 5 Then az = 1
                            If TMPTmpCutGridX - az < 0 Then az = 0
                    
                            ObjMe.MousePointer = 2
                            XIndexGrid = GridX(HeadSellOnFixCol(TMPTmpCutGridX - az)).GridRealOnPosisi
                            XPointerIndex = HeadSellOnFixCol(TMPTmpCutGridX - az)
                    Else
                        ObjMe.MousePointer = 0
                    End If
                Exit Do
                End If
                TMPTmpCutGridX = TMPTmpCutGridX + 1
                If TMPTmpCutGridX = ShowOnConHeadSell Then
                    MyBlock = False
                    TmpCutGridX = TmpCutGridX + 1
                End If
            End If
'            >>>>>>>>>>>>>>>>>>>>>>> >>>>>>>> XXXX
            If MyBlock = False And TmpCutGridX < SellCountColumn Then
                UpBlock = True
                Cx = SellLeft(ShowFixCol(TmpCutGridX)) + SellWidth(ShowFixCol(TmpCutGridX))
                If X < Cx Or TmpCutGridX > SellCountColumn - 1 Then
                    If X > Cx - 5 Or X < SellLeft(ShowFixCol(TmpCutGridX)) + 5 Then
                        If X < SellLeft(ShowFixCol(TmpCutGridX)) + 5 Then az = 1
                            If TmpCutGridX - az < 0 Then az = 0
                            
                            ObjMe.MousePointer = 2
                            XIndexGrid = (TmpCutGridX - az)
                            XPointerIndex = ShowFixCol(TmpCutGridX - az)
                            t = t
                    Else
                        ObjMe.MousePointer = 0
                    End If
                Exit Do
                End If
            End If
        Else
            Exit Do
        End If
        TmpCutGridX = TmpCutGridX + 1
    Loop
'Else
'    If ObjMe2.Visible = False Then ObjMe.MousePointer = 0
'End If
Command21.Caption = XIndexGrid & " >>> " & XPointerIndex
'MsgBox ShowFixCol(TmpCutGridX)
End Sub

'Private Sub Picture Events -------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub MouseX(tBGrid As Integer, ObjMe As Object, ObjMe2 As Object, X As Single, Y As Single, MdlCountX As Integer)
Dim TMPTmpCutGridX As Integer, TmpCutGridX As Integer
Dim MyBlock As Boolean, UpBlock As Boolean

'If Y < FixedHeight Then
    TmpCutGridX = MdlCountX
    TMPTmpCutGridX = HscHead.Value
    Do
        If TmpCutGridX < SellCountColumn - HideCountFixCol Then
            'Bisa juga di gunakan dari jumlah lebar pada headnya
            If UpBlock = False And ShowOnConHeadSell > 0 Then
                MyBlock = True
                Cx = SellLeft(HeadSellOnFixCol(TMPTmpCutGridX)) + SellWidth(HeadSellOnFixCol(TMPTmpCutGridX))
                If X < Cx Or TmpCutGridX > SellCountColumn - 1 Then
                    If X > Cx - 5 Or X < SellLeft(HeadSellOnFixCol(TMPTmpCutGridX)) + 5 Then
                        If X < SellLeft(HeadSellOnFixCol(TMPTmpCutGridX)) + 5 Then az = 1
                            If TMPTmpCutGridX - az < 0 Then az = 0
                    
                            ObjMe.MousePointer = 2
                            XIndexGrid = GridX(HeadSellOnFixCol(TMPTmpCutGridX - az)).GridRealOnPosisi
                            XPointerIndex = HeadSellOnFixCol(TMPTmpCutGridX - az)
                    Else
                        ObjMe.MousePointer = 0
                    End If
                Exit Do
                End If
                TMPTmpCutGridX = TMPTmpCutGridX + 1
                If TMPTmpCutGridX = ShowOnConHeadSell + HscHead.Value Then
                    MyBlock = False
                    TmpCutGridX = TmpCutGridX + 1
                End If
            End If
'            >>>>>>>>>>>>>>>>>>>>>>> >>>>>>>> XXXX
            If MyBlock = False And TmpCutGridX < SellCountColumn Then
                UpBlock = True
                Cx = SellLeft(ShowFixCol(TmpCutGridX)) + SellWidth(ShowFixCol(TmpCutGridX))
                If X < Cx Or TmpCutGridX > SellCountColumn - 1 Then
                    If X > Cx - 5 Or X < SellLeft(ShowFixCol(TmpCutGridX)) + 5 Then
                        If X < SellLeft(ShowFixCol(TmpCutGridX)) + 5 Then az = 1
                            If TmpCutGridX - az < 0 Then az = 0
                            
                            ObjMe.MousePointer = 2
                            XIndexGrid = (TmpCutGridX - az)
                            XPointerIndex = ShowFixCol(TmpCutGridX - az)
                            t = t
                    Else
                        ObjMe.MousePointer = 0
                    End If
                Exit Do
                End If
            End If
        Else
            Exit Do
        End If
        TmpCutGridX = TmpCutGridX + 1
    Loop
'Else
'    If ObjMe2.Visible = False Then ObjMe.MousePointer = 0
'End If
Command21.Caption = XIndexGrid & " >>> " & XPointerIndex
'MsgBox ShowFixCol(TmpCutGridX)
End Sub
Sub MouseUpX(tBGrid As Integer, ObjMe1 As Object, ObjMe2 As Object)
Dim TmpGridWidth As Integer
    
    Command21.Caption = XIndexGrid & " - " & XPointerIndex
    If ObjMe1.MousePointer = 2 Then
        ObjMe2.Visible = False
        ObjMe1.MousePointer = 0
        
        If SellWidth(ShowFixCol(XIndexGrid)) = 0 And XMovGrid > 0 Then XIndexGrid = XIndexGrid - 1
        
        TmpGridWidth = (SellWidth(ShowFixCol(XIndexGrid))) + -XMovGrid
        If TmpGridWidth < 0 Then TmpGridWidth = 0
    
        If TmpGridWidth > SellWidth_Def - 1 Then
            SellWidth(ShowFixCol(XIndexGrid), True) = TmpGridWidth
        Else
            SellWidth(ShowFixCol(XIndexGrid), True) = SellWidth_Def
        End If
    End If
    
End Sub

Sub MouseY(tBGrid As Integer, ObjMe As Object, ObjMe2 As Object, X As Single, Y As Single, MdlCountY As Integer)
Dim TmpCutGridY As Integer

'If X < FixedWidth Then
    TmpCutGridY = MdlCountY
    Do
        If TmpCutGridY < SellCountRow Then
            Cy = SellTop(TmpCutGridY) + SellHeight(TmpCutGridY)
            If Y < Cy Or TmpCutGridY > SellCountRow - 1 Then
                If Y > Cy - 5 Or Y < SellTop(TmpCutGridY) + 5 Then
                    If Y < SellTop(TmpCutGridY) + 5 Then az = 1
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
Sub MouseUpY(tBGrid As Integer, ObjMe1 As Object, ObjMe2 As Object)
Dim TmpGridHeight As Integer
    
    If ObjMe1.MousePointer = 2 Then
        ObjMe2.Visible = False
        ObjMe1.MousePointer = 0
        
        If SellHeight(YIndexGrid) = 0 And YMovGrid > 0 Then YIndexGrid = YIndexGrid - 1
        
        TmpGridHeight = (SellHeight(YIndexGrid) - SellHeight_Def) + -YMovGrid
        If TmpGridHeight < 0 Then TmpGridHeight = 0
    
        If TmpGridHeight > (SellHeight_Def - SellHeight_Def) - 1 Then _
            SellHeight(YIndexGrid, True) = TmpGridHeight Else _
            SellHeight(YIndexGrid, True) = SellHeight_Def
    End If
End Sub

Function SizeY(tBGrid As Integer, X As Single, Y As Single) As Boolean
    If X > RangePicSubX1 And X < RangePicSubX1 + RangePicSubX2 And _
    Y > SellTop(YPointerIndex) + RangePicSubY1 And Y < SellTop(YPointerIndex) + RangePicSubY1 + RangePicSubY2 Then
    SizeY = True
        If SellHeight(YPointerIndex) > SellHeight_Def Then
            SizeUpDown YPointerIndex, False
        Else
            SizeUpDown YPointerIndex, True
        End If
    End If
End Function


Sub SearchPointer(tBGrid As Integer, X As Single, TmpCutGridX As Long, Y As Single, TmpCutGridY As Long)
Dim TMPTmpCutGridX As Integer
Dim MyBlock As Boolean, UpBlock As Boolean
Dim IndexPx As Integer
Dim XEndGrid As Integer
         'TmpCutGridX = Form1.Hsc.Value
'         TmpCutGridX = 0 ShowFixCol
'>>>>>>>>>>>
    'If ShowOnConHeadSell - 1 + Hsc.Value >= SellCountColumn - 1 Then MsgBox ""
If ShowOnConHeadSell - 1 + Hsc.Value < SellCountColumn - 1 Then
    IndexPx = ShowFixCol(GridRightX)
    XEndGrid = SellLeft(IndexPx) + SellWidth(IndexPx)
Else
    IndexPx = HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))
    XEndGrid = (SellLeft(IndexPx)) + (SellWidth(IndexPx)) - ((SellLeft(IndexPx)) + (SellWidth(IndexPx)) - RangeEndHead)
End If

'If X > (SellLeft(IndexPx)) + (SellWidth(IndexPx)) - ((SellLeft(IndexPx)) + (SellWidth(IndexPx)) - RangeEndHead) Then MsgBox "Ok"

If X < XEndGrid Then 'SellLeft(IndexPx) + SellWidth(IndexPx) Then  'MsgBox ""
    If X > RangeEndHead And ShowOnConHeadSell > 0 Then
        UpBlock = True
        MyBlock = False
        If ShowOnConHeadSell - 1 + Hsc.Value < SellCountColumn - 1 Then TmpCutGridX = TmpCutGridX + ShowOnConHeadSell
        '^----------> Error
    End If
    Do '--> X
        If UpBlock = False And ShowOnConHeadSell > 0 Then
'If SellLeft(HeadSellOnFixCol(TMPTmpCutGridX)) + SellWidth(HeadSellOnFixCol(TMPTmpCutGridX)) > RangeEndHead Then MsgBox TMPTmpCutGridX
            MyBlock = True
            If X < SellLeft(HeadSellOnFixCol(TMPTmpCutGridX + HscHead.Value)) + SellWidth(HeadSellOnFixCol(TMPTmpCutGridX + HscHead.Value)) _
            Or TmpCutGridX >= GridRightX Then
                TmpCutGridX = -(TMPTmpCutGridX + 1 + HscHead.Value)
                Exit Do
            End If
            TMPTmpCutGridX = TMPTmpCutGridX + 1
                If TMPTmpCutGridX = ShowOnConHeadSell Then
                    MyBlock = False
                    TmpCutGridX = TmpCutGridX + 1
                End If
        End If
        If MyBlock = False Then
            UpBlock = True
            If X < SellLeft(ShowFixCol(TmpCutGridX)) + SellWidth(ShowFixCol(TmpCutGridX)) _
            Or TmpCutGridX >= GridRightX Then Exit Do
        End If
    TmpCutGridX = TmpCutGridX + 1
    Loop
Else
    If ShowOnConHeadSell - 1 + Hsc.Value < SellCountColumn - 1 Then
        TmpCutGridX = GridRightX
    Else
        TmpCutGridX = -(ShowOnConHeadSell + HscHead.Value)
    End If
End If
    'TmpCutGridX = TmpCutGridX + HscHead.Value
    Command24.Caption = TmpCutGridX & " Click**"
    'TmpCutGridX = ShowFixCol(TmpCutGridX)
   ' MsgBox X & " " & SellLeft(0) + SellWidth(0) & " " & TmpCutGridX
        'TmpCutGridY = Form1.Vsc.Value
    'nForm(0).Text3.SellText = ""
    Do '--> Y
        If Y < SellTop(TmpCutGridY) + SellHeight(TmpCutGridY) _
        Or TmpCutGridY >= GridDownY Then Exit Do
        'nForm(0).Text3.SellText = "nForm(0).Text3.SellText" & TmpCutGridY & "." & vbCrLf
    TmpCutGridY = TmpCutGridY + 1
    Loop
End Sub

Sub TMPSearchPointer(tBGrid As Integer, X As Single, TmpCutGridX As Long, Y As Single, TmpCutGridY As Long)
         'TmpCutGridX = Form1.Hsc.Value
'         TmpCutGridX = 0 ShowFixCol
'>>>>>>>>>>>
    Do
         If X < SellLeft(ShowFixCol(TmpCutGridX)) + SellWidth(ShowFixCol(TmpCutGridX)) _
         Or TmpCutGridX >= GridRightX Then Exit Do
    TmpCutGridX = TmpCutGridX + 1
    Loop
    'TmpCutGridX = ShowFixCol(TmpCutGridX)
   ' MsgBox X & " " & SellLeft(0) + SellWidth(0) & " " & TmpCutGridX
        'TmpCutGridY = Form1.Vsc.Value
    'nForm(0).Text3.SellText = ""
    Do
         If Y < SellTop(TmpCutGridY) + SellHeight(TmpCutGridY) _
         Or TmpCutGridY >= GridDownY Then Exit Do
        'nForm(0).Text3.SellText = "nForm(0).Text3.SellText" & TmpCutGridY & "." & vbCrLf
    TmpCutGridY = TmpCutGridY + 1
    Loop
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub PointerGrid_Set(nPXY As PointInvert)
    Pxy = nPXY
    CountTpXy_X = Pxy.Px2 - Pxy.Px1
    CountTpXy_Y = Pxy.Py2 - Pxy.Py1
    
    DrawInvertToGrid CBGrid, Picture1, Pxy.Px1, Pxy.Py1, Pxy.Px2, Pxy.Py2
End Sub
Function PointerGrid_Get() As PointInvert
    PointerGrid_Get = Pxy
End Function


'Invert Paint-------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub C___DrawInvertToGrid(ByVal tBGrid As Integer, ByVal ObjMe As Object, iPxy As PointInvert, ByVal Px1 As Long, ByVal Py1 As Long, ByVal Px2 As Long, ByVal Py2 As Long)
Dim TmpPxy As PointInvert
Dim Cx As Long, Cy As Long


If (Px1 > SetNewGrid.GridXCount Or Px2 > SetNewGrid.GridXCount) Or _
   (Py1 > SetNewGrid.GridYCount Or Py2 > SetNewGrid.GridYCount) Then
    Px1 = 0
    Py1 = 0
    Px2 = 0
    Py2 = 0
End If
'Command19.Caption = SetNewGrid.GridXCount & " " & SetNewGrid.GridYCount


' And Hsc.Value < iPxy.Px2 + 1
If (GridRightX > iPxy.Px1 - 1 And Hsc.Value < iPxy.Px2 + 1) And _
GridDownY > iPxy.Py1 - 1 And Vsc.Value < iPxy.Py2 + 1 Then

iPxy.Px1 = Px1
iPxy.Py1 = Py1
iPxy.Px2 = Px2
iPxy.Py2 = Py2

'    If iPxy.Px2 >= SellCountColumn Then
'        iPxy.Px1 = SellCountColumn - 1
'        iPxy.Px2 = SellCountColumn - 1
'    End If
'    If iPxy.Py2 >= SellCountRow Then
'        iPxy.Py1 = SellCountRow - 1
'        iPxy.Py2 = SellCountRow - 1
'    End If
    If SellLeft(Px1) = 0 Then Px1 = GridLeftX
    If SellTop(Py1) = 0 Then Py1 = GridUpY

    If GridLeftX > Px1 Then Px1 = GridLeftX
    If Px2 > GridRightX Then Px2 = GridRightX
    
    If GridUpY > Py1 Then Py1 = GridUpY
    If Py2 > GridDownY Then Py2 = GridDownY
    
    Cx = SellLeft(Px2) + SellWidth(Px2) - 1
    Cy = SellTop(Py2) + SellHeight(Py2) - 1
    
    If Cx = 0 Then
        Cx = SellLeft(GridRightX) + SellWidth(GridRightX)
    End If
    If Cy = 0 Then
        Cy = SellTop(GridDownY) + SellHeight(GridDownY)
    End If
    
    'DrawInvert ObjMe, SellLeft(Px1) + 2, SellTop(Py1) + 2, Cx, Cy
    Shape1.left = SellLeft(Px1) + 2
    Shape1.top = SellTop(Py1) + 2
    Shape1.Width = SellWidth(Px2) - 3
    Shape1.Height = 25 - 3 'SellHeight(Py2) - 3
'    DrawInvert ObjMe, SellLeft(XPointerIndex) + 3, SellTop(YPointerIndex + 2) + 3, _
    Cx - 2, Cy - 2
'End If
End If

Command3.Caption = SellLeft(Px1) & " " & Px2 & " " & Cx
End Sub

'Invert Paint-------------------------------------------------------------------------------------------------------------------------------------------------------------
'Private Sub DrawInvertToGrid(ByVal tBGrid As Integer, ByVal ObjMe As Object, iPxy As PointInvert, ByVal Px1 As Long, ByVal Py1 As Long, ByVal Px2 As Long, ByVal Py2 As Long)
Private Sub DrawInvertToGrid(ByVal tBGrid As Integer, ByVal ObjMe As Object, ByVal Px1 As Long, ByVal Py1 As Long, ByVal Px2 As Long, ByVal Py2 As Long)
'Px1 & Px2 Output Real FixedX

Dim TmpPxy As PointInvert
Dim Cx As Long, Cy As Long
Dim iPxy As PointInvert, iX As Integer, OpenProsesi As Boolean, Px1Head As Integer, Px2Head As Integer
Dim Wids As Integer
Dim TsX1n As Integer, TsX2n As Integer
Dim TsY1n As Integer, TsY2n As Integer

Command28.Caption = Px1 & " " & ShowFixCol(Px1) & " " & GridX(ShowFixCol(Px1)).GridValue

If Px1 > SellCountColumn - 1 Or Px2 > SellCountColumn - 1 Then MsgBox "Error In Nomber 1"

Command35.Caption = Pxy.Px1 & " " & Pxy.Px2

Command33.Caption = ""
'Exit Sub
'px1 in eror jadi di rubah

iPxy.Px1 = Px1
iPxy.Py1 = Py1
iPxy.Px2 = Px2
iPxy.Py2 = Py2

'On Error GoTo Errors
Command25.Caption = Px1
If 8 = 9 And (Px1 > SetNewGrid.GridXCount Or Px2 > SetNewGrid.GridXCount) Or _
   (Py1 > SetNewGrid.GridYCount Or Py2 > SetNewGrid.GridYCount) Then
    Px1 = 0
    Py1 = 0
    Px2 = 0
    Py2 = 0
End If

Dim GoOpen As Boolean

Command26.Caption = ShowOnConHeadSell 'GridX(Px1).GridIndexHead

If Px1 < 0 Then Px1 = HeadSellOnFixCol(Abs(Px1) - 1)
If Px2 < 0 Then Px2 = HeadSellOnFixCol(Abs(Px2) - 1)


If Px1 > SellCountColumn - 1 Or Px2 > SellCountColumn - 1 Then MsgBox "Error In Nomber 2"

'Jika Untuk Proses Dari Code Maka Ditetapkan Variabel Dari Proses Code
'
'MsgBox GridX(GridX(ShowFixCol(Px1 + TPxy.Px1)).GridRealPosisi).GridIndexHead
If ((ShowOnConHeadSell > 0 And GridX(ShowFixCol((Px1 + TPxy.Px1))).GridIndexHead > 0) Or (GridRightX > Px1 + TPxy.Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px1 + TPxy.Px1 + 1)) And _
   (GridDownY > Py1 + TPxy.Py1 - 1 And Vsc.Value < Py1 + TPxy.Py1 + 1) Then
    
    If ShowOnConHeadSell > 0 Then
        If GridX(ShowFixCol(Px1 + TPxy.Px1)).GridRealPosisi <= GridX(ShowFixCol(Hsc.Value + (ShowOnConHeadSell - 1))).GridRealPosisi And _
          (GridX(ShowFixCol(Px1 + TPxy.Px1)).GridRealPosisi < GridX(HeadSellOnFixCol(HscHead.Value)).GridRealPosisi Or GridX(ShowFixCol(Px1 + TPxy.Px1)).GridRealPosisi > GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealPosisi) Then
            OpenProsesi = True
        End If
    End If
    
    If OpenProsesi = False Then
'        Shape1.left = SellLeft(GridX(ShowFixCol(Px1 + TPxy.Px1)).GridRealPosisi) + 2
        If ShowOnConHeadSell > 0 Then
            Shape1.left = SellLeft(ShowFixCol(Px1 + TPxy.Px1)) + 2 '  . . . . Error Posision
        Else
            Shape1.left = SellLeft(ShowFixCol(Px1 + TPxy.Px1)) + 2 ' . . . . Error Posision
        End If
        Shape1.top = SellTop(Py1 + TPxy.Py1) + 2
        If SellWidth(ShowFixCol(Px1)) <= 0 Then SellWidth(ShowFixCol(Px1)) = 3
        Shape1.Width = SellWidth(ShowFixCol(Px1 + TPxy.Px1)) - 3  'Abs(SellLeft(Px1) - SellLeft(Px2)) + SellWidth(Px2) - 3
        If SellHeight(Py1) <= 0 Then SellHeight(Py1) = 3
        Shape1.Height = SellHeight(Py1 + TPxy.Py1) - 3 'Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3 '25 - 3 'SellHeight(Py2) - 3
        
        If Shape1.Visible = False Then Shape1.Visible = True
    Else
        If Shape1.Visible = True Then Shape1.Visible = False
        Command33.Caption = "L 1"
    End If
Else
    If Shape1.Visible = True Then Shape1.Visible = False
    Command33.Caption = "L 2"
End If
OpenProsesi = False
    
Shape3.left = Shape1.left - 1
Shape3.top = Shape1.top - 1
Shape3.Width = Shape1.Width + 2
Shape3.Height = Shape1.Height + 2
Shape3.Visible = Shape1.Visible

Command30.Caption = TPxy.Px1 & " " & TPxy.Px2 & " - " & TPxy.Py1 & " " & TPxy.Py2

Px1Head = -1
If 8 = 8 And ShowOnConHeadSell > 0 And Hsc.Value + ShowOnConHeadSell <= SellCountColumn Then
    
    If (((Px1) <= GridX(ShowFixCol(Hsc.Value + ShowOnConHeadSell - 1)).GridRealOnPosisi And GridX(ShowFixCol(Px1)).GridIndexHead = 0) Or ((Px2) <= GridX(ShowFixCol(Hsc.Value + ShowOnConHeadSell - 1)).GridRealPosisi And GridX(ShowFixCol(Px2)).GridIndexHead = 0)) Then
        For iX = HscHead.Value To HscHead.Value + (ShowOnConHeadSell - 1)
            If (Px1) <= GridX(HeadSellOnFixCol(iX)).GridRealPosisi And (Px2) >= GridX(HeadSellOnFixCol(iX)).GridRealPosisi Then
                If Px1Head = -1 Then Px1Head = iX: Px2Head = iX Else Px2Head = iX
                OpenProsesi = True
                    If Px2 <= GridX(HeadSellOnFixCol(iX)).GridRealPosisi Then Exit For
            End If
        Next iX
    End If
End If
'If GridX(Px1).GridRealOnPosisi = -1 Then GridX(Px1).GridRealOnPosisi = Px1
'If GridX(Px2).GridRealOnPosisi = -1 Then GridX(Px2).GridRealOnPosisi = Px2

If GridX(ShowFixCol(Px1)).GridRealOnPosisi = -1 Then GridX(ShowFixCol(Px1)).GridRealOnPosisi = Px1
If GridX(ShowFixCol(Px2)).GridRealOnPosisi = -1 Then GridX(ShowFixCol(Px2)).GridRealOnPosisi = Px2

If ShowOnConHeadSell > 0 Then
    If ((OpenProsesi = True Or _
    (GridX(ShowFixCol(Px1)).GridIndexHead > 0 And Px1 >= GridX(HeadSellOnFixCol(HscHead.Value)).GridRealPosisi And Px1 <= GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealPosisi) Or _
    (GridX(ShowFixCol(Px2)).GridIndexHead > 0 And Px2 >= GridX(HeadSellOnFixCol(HscHead.Value)).GridRealPosisi And Px2 <= GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealPosisi)) Or _
    (GridRightX > Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px2 + 1)) And _
    (GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1) Then

        If Px2 <> Px1 Or Py2 <> Py1 Then
            If (Px1) > GridX(HeadSellOnFixCol((ShowOnConHeadSell - 1) + HscHead.Value)).GridRealPosisi Then
                If Px1 < ShowOnConHeadSell + Hsc.Value Then TsX1n = IndexEndHead Else TsX1n = ShowFixCol(Px1)
            Else
                If GridX(ShowFixCol(Px1)).GridRealOnPosisi < GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi And Px2 <= GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi Then
                    TsX1n = GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi
                Else
                    If GridX(ShowFixCol(Px1)).GridIndexHead = 0 And Px1Head > -1 Then TsX1n = (HeadSellOnFixCol(Px1Head)) Else TsX1n = ShowFixCol(Px1) 'Else
                End If
            End If
            Shape2.left = SellLeft(TsX1n) + 2
            
            If Px2 >= ShowOnConHeadSell + Hsc.Value Then
                If Px2 <= GridRightX Then TsX2n = ShowFixCol(Px2) Else TsX2n = ShowFixCol(GridRightX)
            Else
    '            If Px1Head > -1 Then
                    If GridX(ShowFixCol(Px1)).GridIndexHead = 0 Then TsX1n = HeadSellOnFixCol(Px1Head) Else TsX1n = ShowFixCol(Px1)
                    If GridX(ShowFixCol(Px2)).GridIndexHead = 0 Then TsX2n = HeadSellOnFixCol(Px2Head) Else TsX2n = ShowFixCol(Px2)
                    
                    If GridX(TsX2n).GridIndexHead > ShowOnConHeadSell + HscHead.Value Then
                        TsX2n = HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))
                    End If
                    
                    If SellLeft(TsX2n) + SellWidth(TsX2n) > RangeEndHead And RangeEndHead > 0 Then
                        Wids = SellLeft(TsX2n) + SellWidth(TsX2n) - RangeEndHead + 2
                    End If
    '            End If
            End If
            Shape2.Width = Abs(SellLeft(TsX1n) - SellLeft(TsX2n)) + SellWidth(TsX2n) - Wids - 2
                
            If Py1 >= Vsc.Value Then TsY1n = Py1 Else TsY1n = Vsc.Value
            Shape2.top = SellTop(TsY1n) + 2
            
            If Py2 <= GridDownY Then TsY2n = Py2 Else TsY2n = GridDownY
            Shape2.Height = Abs(SellTop(TsY1n) - SellTop(TsY2n)) + SellHeight(TsY2n) - 2
            
            If Shape2.Visible = False Then Shape2.Visible = True
        Else
            If Shape2.Visible = True Then Shape2.Visible = False
        End If
    Else
        If Shape2.Visible = True Then Shape2.Visible = False
    End If
Else
    If (GridRightX > Px1 - 1 And Hsc.Value < Px2 + 1) And (GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1) Then
        If Px2 <> Px1 Or Py2 <> Py1 Then
            If Px1 >= Hsc.Value Then TsX1n = ShowFixCol(Px1) Else TsX1n = ShowFixCol(Hsc.Value)
            Shape2.left = SellLeft(TsX1n) + 2
            
            If Px2 <= GridRightX Then TsX2n = ShowFixCol(Px2) Else TsX2n = ShowFixCol(GridRightX)
            Shape2.Width = Abs(SellLeft(TsX1n) - SellLeft(TsX2n)) + SellWidth(TsX2n) - Wids - 2
            
            If Py1 >= Vsc.Value Then TsY1n = Py1 Else TsY1n = Vsc.Value
            Shape2.top = SellTop(TsY1n) + 2
            
            If Py2 <= GridDownY Then TsY2n = Py2 Else TsY2n = GridDownY
            Shape2.Height = Abs(SellTop(TsY1n) - SellTop(TsY2n)) + SellHeight(TsY2n) - 2
            
            If Shape2.Visible = False Then Shape2.Visible = True
        Else
            If Shape2.Visible = True Then Shape2.Visible = False
            Command34.Caption = "M 1"
        End If
    Else
        If Shape2.Visible = True Then Shape2.Visible = False
        Command34.Caption = "M 2"
    End If
End If

Shape4.left = Shape2.left '+ 2
Shape4.top = Shape2.top '+ 2
Shape4.Width = Shape2.Width '- 4
Shape4.Height = Shape2.Height '- 4
Shape4.Visible = Shape2.Visible

Line1.X1 = Shape2.left - 5
Line1.X2 = Shape2.left - 5
Line1.Y1 = Shape2.top - 3
Line1.Y2 = Line1.Y1 + 20
If Hsc.Value > Px1 - 0 Then Line1.Visible = False Else Line1.Visible = Shape2.Visible

Line2.X1 = Shape2.left - 3
Line2.X2 = Line2.X1 + 20
Line2.Y1 = Shape2.top - 5
Line2.Y2 = Shape2.top - 5
If Vsc.Value > Py1 - 0 Then Line2.Visible = False Else Line2.Visible = Shape2.Visible

Shape5.left = Shape2.left - Shape5.Width - 3
Shape5.top = Shape2.top - Shape5.Height - 3
If Line1.Visible = False And Line2.Visible = False Then Shape5.Visible = False Else Shape5.Visible = Shape2.Visible


Command36.Caption = Pxy.Px1 & " " & Pxy.Px2


Pxy = iPxy

'Pxy.Px1 = Px1
'Pxy.Py1 = Py1
'Pxy.Px2 = Px2
'Pxy.Py2 = Py2
Exit Sub
100
If OpenProsesi = True Or _
((GridX(ShowFixCol(Px1)).GridIndexHead > 0 And Px1 > HscHead.Value + ShowOnConHeadSell) Or _
(GridX(ShowFixCol(Px2)).GridIndexHead > 0 And Px2 > HscHead.Value + ShowOnConHeadSell)) Or _
(((GridRightX > Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px2 + 1) And _
GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1)) Then

Command25.Caption = ShowOnConHeadSell

iPxy.Px1 = Px1: iPxy.Py1 = Py1: iPxy.Px2 = Px2: iPxy.Py2 = Py2

    Command30.Caption = ""
    If SellLeft(ShowFixCol(Px1)) = 0 Then
    Px1 = GridLeftX
    Command30.Caption = "OK"
    End If
    If SellTop(Py1) = 0 Then Py1 = GridUpY

    If GridUpY > Py1 Then Py1 = GridUpY
    If Py2 > GridDownY Then Py2 = GridDownY
    
    Cx = SellLeft(ShowFixCol(Px2)) + SellWidth(ShowFixCol(Px2)) - 1
    Cy = SellTop(Py2) + SellHeight(Py2) - 1
        If Cx = 0 Then Cx = SellLeft(ShowFixCol(GridRightX)) + SellWidth(ShowFixCol(GridRightX))
        If Cy = 0 Then Cy = SellTop(GridDownY) + SellHeight(GridDownY)
    
    '|-> Jika Terjadi Pergeseran Pada HSC Atau VSC Secara Back -------------------------------------------------------------------------
    If Hsc.Value > Px1 And GridX(ShowFixCol(Px1)).GridIndexHead = 0 Then GridX(ShowFixCol(Px1)).GridLeft = SetNewGrid.GridSize.GDRangeX
    If Px2 > GridRightX And GridX(ShowFixCol(Px2)).GridIndexHead = 0 Then GridX(ShowFixCol(Px2)).GridLeft = Picture1.ScaleWidth
    If Vsc.Value > Py1 Then GridY(Py1).GridTop = SetNewGrid.GridSize.GDRangeY
    If Py2 > GridDownY Then GridY(Px2).GridHeight = Picture1.ScaleHeight
    '-----------------------------------------------------------------------------------------------------------------------------------
    
    If ShowOnConHeadSell > 0 And Hsc.Value >= Px1 And GridX(ShowFixCol(Px1)).GridIndexHead = 0 Then
        GridX(ShowFixCol(Px1)).GridLeft = GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth + 4
    Else
'        GridX(ShowFixCol(Px1)).GridLeft = GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth + 4
    End If
    
    Command25.Caption = Px1 & " " & ShowFixCol(Px1) & " " & GridX(Px1).GridRealOnPosisi
    Shape1.left = SellLeft(ShowFixCol(Px1)) + 2
    Shape1.top = SellTop(Py1) + 2
    If SellWidth(ShowFixCol(Px1)) <= 0 Then SellWidth(ShowFixCol(Px1)) = 3
    Shape1.Width = SellWidth(ShowFixCol(Px1)) - 3 'Abs(SellLeft(Px1) - SellLeft(Px2)) + SellWidth(Px2) - 3
    If SellHeight(Py1) <= 0 Then SellHeight(Py1) = 3
    Shape1.Height = SellHeight(Py1) - 3 'Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3 '25 - 3 'SellHeight(Py2) - 3
    Shape1.Visible = True
    
    If Px2 <> Px1 Or Py2 <> Py1 Then
        'If SellLeft(ShowFixCol(Px1)) <> 25 Then
'        Shape2.left = SellLeft(ShowFixCol(Px1)) + 2
        
If 8 = 9 Then
'        MsgBox Px1Head & " " & Px2Head

'        If ShowOnConHeadSell = 0 And Px1 >= GridLeftX + ShowOnConHeadSell Then
        If ShowOnConHeadSell = 0 Then
            TsX1n = ShowFixCol(Px1)
        Else
'            TsX1n = ShowFixCol(Px1)
'            TsX1n = HeadSellOnFixCol(Px1Head)
            If GridX(Px1).GridIndexHead = 0 Then TsX1n = HeadSellOnFixCol(Px1Head) Else TsX1n = ShowFixCol(Px1) 'Else
        End If
        Shape2.left = SellLeft(TsX1n) + 2
        
        If ShowOnConHeadSell = 0 And Px1 >= GridLeftX + ShowOnConHeadSell Then
            TsX1n = ShowFixCol(Px1)
            TsX2n = ShowFixCol(Px2)
        Else
            If Px1 < ShowOnConHeadSell + Hsc.Value And Px2 < ShowOnConHeadSell + Hsc.Value Then
'                TsX1n = HeadSellOnFixCol(Px1Head)
'                TsX1n = ShowFixCol(Px1)
                If GridX(Px1).GridIndexHead = 0 Then TsX1n = HeadSellOnFixCol(Px1Head) Else TsX1n = ShowFixCol(Px1) 'Else
                
                If GridX(Px2).GridIndexHead = 0 Then TsX2n = HeadSellOnFixCol(Px2Head) Else TsX2n = ShowFixCol(Px2) 'Else
                If SellLeft(HeadSellOnFixCol(Px2Head)) + SellWidth(HeadSellOnFixCol(Px2Head)) > RangeEndHead And RangeEndHead > 0 Then
                    Wids = SellLeft(HeadSellOnFixCol(Px2Head)) + SellWidth(HeadSellOnFixCol(Px2Head)) - RangeEndHead
                End If
            Else
            
            End If
        End If
        
        Shape2.Width = Abs(SellLeft(TsX1n) - SellLeft(TsX2n)) + SellWidth(TsX2n) - Wids
        Shape2.top = SellTop(Py1) + 2
        Shape2.Height = Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3
    If Shape2.Visible = False Then Shape2.Visible = True
End If
        
If 8 = 8 Then
        If ShowOnConHeadSell > 0 And OpenProsesi = True Then
            Shape2.left = SellLeft(HeadSellOnFixCol(Px1Head))
            If Px2 > Hsc.Value + Px2Head Then
                Shape2.Width = Abs(SellLeft(HeadSellOnFixCol(Px1Head)) - SellLeft(ShowFixCol(Px2))) + SellWidth(ShowFixCol(Px2)) '- 25 'SellLeft(ShowFixCol(Px2)) '+ SellWidth(ShowFixCol(Px2))
            Else
                If SellLeft(HeadSellOnFixCol(Px2Head)) + SellWidth(HeadSellOnFixCol(Px2Head)) > RangeEndHead And RangeEndHead > 0 Then
                    Wids = SellLeft(HeadSellOnFixCol(Px2Head)) + SellWidth(HeadSellOnFixCol(Px2Head)) - RangeEndHead
                End If
                Shape2.Width = Abs(SellLeft(HeadSellOnFixCol(Px1Head)) - SellLeft(HeadSellOnFixCol(Px2Head))) + SellWidth(HeadSellOnFixCol(Px2Head)) - Wids
            End If
        Else
'            If ShowOnConHeadSell > 0 And Px1 < Hsc.Value + ShowOnConHeadSell And GridX(ShowFixCol(Px1)).GridIndexHead = 0 Then
            If ShowOnConHeadSell > 0 And Px1 > Hsc.Value + ShowOnConHeadSell Then
'                Shape2.left = SellLeft(HeadSellOnFixCol(ShowOnConHeadSell - 1)) + SellWidth(HeadSellOnFixCol(ShowOnConHeadSell - 1)) '+ 3 + 3
'                Shape2.left = SellLeft(ShowFixCol(ShowOnConHeadSell + GridLeftX)) '+ SellWidth(ShowFixCol(ShowOnConHeadSell + GridLeftX)) '+ 3 + 3
                'Shape2.left = SellLeft(10) + 3
'                MsgBox ShowOnConHeadSell + GridLeftX
            End If
            
            If ShowOnConHeadSell > 0 Then
                If ShowFixCol(Px2) = HeadSellOnFixCol(ShowOnConHeadSell - 1) And RangeEndHead > 0 Then 'MsgBox ""
                    If SellLeft(ShowFixCol(Px2)) + SellWidth(ShowFixCol(Px2)) > RangeEndHead And GridX(Px2).GridIndexHead > 0 Then
                        Wids = SellLeft(ShowFixCol(Px2)) + SellWidth(ShowFixCol(Px2)) - RangeEndHead - 1
                    End If
                End If
            End If
            If ShowFixCol(Px1) > HeadSellOnFixCol(ShowOnConHeadSell - 1) And ShowFixCol(Px1) < ShowFixCol(ShowOnConHeadSell + GridLeftX) Then
                Shape2.left = SellLeft(ShowFixCol(ShowOnConHeadSell + GridLeftX)) '+ SellWidth(ShowFixCol(ShowOnConHeadSell + GridLeftX)) '+ 3 + 3
                Shape2.Width = Abs(SellLeft(ShowFixCol(ShowOnConHeadSell + GridLeftX)) - SellLeft(ShowFixCol(Px2))) + SellWidth(ShowFixCol(Px2))
            Else
                'If Px2 < ShowOnConHeadSell + GridLeftX Then
                Shape2.Width = Abs(SellLeft(ShowFixCol(Px1)) - SellLeft(ShowFixCol(Px2))) + SellWidth(ShowFixCol(Px2)) - 3 - Wids
            End If
        End If
        
        Shape2.Height = Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3 '25 - 3 'SellHeight(Py2) - 3
        If Shape2.Visible = False Then Shape2.Visible = True
End If
    Else
        If Shape2.Visible = True Then Shape2.Visible = False
    End If
Else
    Shape1.Visible = False
    Shape2.Visible = False
End If
'Pusing -Pusing
'GridX(Px1).GridIndexHead = GridX(Px1).GridIndexHead
If ShowOnConHeadSell > 0 And GridX(Px1).GridIndexHead - 1 > 0 And GridX(Px1).GridIndexHead - 1 < ShowOnConHeadSell Then
'    Px1 = -GridX(ShowFixCol(Px1)).GridIndexHead
End If
If ShowOnConHeadSell > 0 And GridX(Px2).GridIndexHead - 1 > 0 And GridX(Px2).GridIndexHead < ShowOnConHeadSell Then
'    Px2 = -GridX(ShowFixCol(Px2)).GridIndexHead
End If

Pxy.Px1 = Px1
Pxy.Py1 = Py1
Pxy.Px2 = Px2
Pxy.Py2 = Py2

'Shape1.Refresh
'Shape2.Refresh
Command21.Caption = Px1Head & " - " & Px2Head
'Command3.Caption = SellLeft(Px1) & " " & Px2 & " " & Cx
'clear oce oye
Exit Sub
Errors:
MsgBox "Error"
End Sub
'Invert Paint-------------------------------------------------------------------------------------------------------------------------------------------------------------

'Sub tESTPX(PXc1 As Integer, PXc2 As Integer, PYc1 As Integer, PYc2 As Integer)
'Pxy.Px1 = PXc1
'Pxy.Py1 = PYc1
'Pxy.Px2 = PXc2
'Pxy.Py2 = PYc2
'End Sub

'Invert Paint-------------------------------------------------------------------------------------------------------------------------------------------------------------
'Private Sub DrawInvertToGrid(ByVal tBGrid As Integer, ByVal ObjMe As Object, iPxy As PointInvert, ByVal Px1 As Long, ByVal Py1 As Long, ByVal Px2 As Long, ByVal Py2 As Long)
Private Sub Original_DrawInvertToGrid(ByVal tBGrid As Integer, ByVal ObjMe As Object, ByVal Px1 As Long, ByVal Py1 As Long, ByVal Px2 As Long, ByVal Py2 As Long)
'Px1 & Px2 Output Real FixedX

Dim TmpPxy As PointInvert
Dim Cx As Long, Cy As Long
Dim iPxy As PointInvert, iX As Integer, OpenProsesi As Boolean, Px1Head As Integer, Px2Head As Integer
Dim Wids As Integer
Dim TsX1n As Integer, TsX2n As Integer
Dim TsY1n As Integer, TsY2n As Integer


If Px1 > SellCountColumn - 1 Or Px2 > SellCountColumn - 1 Then MsgBox "Error In Nomber 1"


'Exit Sub
'px1 in eror jadi di rubah

iPxy.Px1 = Px1
iPxy.Py1 = Py1
iPxy.Px2 = Px2
iPxy.Py2 = Py2

'On Error GoTo Errors
Command25.Caption = Px1
If 8 = 9 And (Px1 > SetNewGrid.GridXCount Or Px2 > SetNewGrid.GridXCount) Or _
   (Py1 > SetNewGrid.GridYCount Or Py2 > SetNewGrid.GridYCount) Then
    Px1 = 0
    Py1 = 0
    Px2 = 0
    Py2 = 0
End If

Dim GoOpen As Boolean

Command26.Caption = ShowOnConHeadSell 'GridX(Px1).GridIndexHead

If Px1 < 0 Then Px1 = HeadSellOnFixCol(Abs(Px1) - 1)
If Px2 < 0 Then Px2 = HeadSellOnFixCol(Abs(Px2) - 1)


If Px1 > SellCountColumn - 1 Or Px2 > SellCountColumn - 1 Then MsgBox "Error In Nomber 2"

'Jika Untuk Proses Dari Code Maka Ditetapkan Variabel Dari Proses Code
'

If ((ShowOnConHeadSell > 0 And GridX(ShowFixCol(Px1 + TPxy.Px1)).GridIndexHead > 0) Or (GridRightX > Px1 + TPxy.Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px1 + TPxy.Px1 + 1)) And _
   (GridDownY > Py1 + TPxy.Py1 - 1 And Vsc.Value < Py1 + TPxy.Py1 + 1) Then
    
    If ShowOnConHeadSell > 0 Then
        If Px1 + TPxy.Px1 <= GridX(ShowFixCol(Hsc.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi And _
          (Px1 + TPxy.Px1 < GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi Or Px1 + TPxy.Px1 > GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi) Then
            OpenProsesi = True
        End If
    End If
    
    If OpenProsesi = False Then
        Shape1.left = SellLeft(ShowFixCol(Px1 + TPxy.Px1)) + 2
        Shape1.top = SellTop(Py1 + TPxy.Py1) + 2
        If SellWidth(ShowFixCol(Px1)) <= 0 Then SellWidth(ShowFixCol(Px1)) = 3
        Shape1.Width = SellWidth(ShowFixCol(Px1 + TPxy.Px1)) - 3 'Abs(SellLeft(Px1) - SellLeft(Px2)) + SellWidth(Px2) - 3
        If SellHeight(Py1) <= 0 Then SellHeight(Py1) = 3
        Shape1.Height = SellHeight(Py1 + TPxy.Py1) - 3 'Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3 '25 - 3 'SellHeight(Py2) - 3
        
        If Shape1.Visible = False Then Shape1.Visible = True
    Else
        If Shape1.Visible = True Then Shape1.Visible = False
    End If
Else
    If Shape1.Visible = True Then Shape1.Visible = False
End If
OpenProsesi = False
    
Command30.Caption = TPxy.Px1 & " " & TPxy.Px2 & " - " & TPxy.Py1 & " " & TPxy.Py2
'Dim TPxy As PointInvert
'Dim TTPxy As PointInvert'

'Dim HTPxy As PointInvert



'If 8 = 8 And ShowOnConHeadSell > 0 And (GridX(ShowFixCol(Px1)).GridIndexHead = 0 Or GridX(ShowFixCol(Px2)).GridIndexHead = 0) Then
'tambahkan tmp pada variabel px1 guna pada pergeseran pointer
Px1Head = -1
If 8 = 8 And ShowOnConHeadSell > 0 And Hsc.Value + ShowOnConHeadSell <= SellCountColumn Then
    
    If ((Px1 <= GridX(ShowFixCol(Hsc.Value + ShowOnConHeadSell - 1)).GridRealOnPosisi And GridX(ShowFixCol(Px1)).GridIndexHead = 0) Or (Px2 <= GridX(ShowFixCol(Hsc.Value + ShowOnConHeadSell - 1)).GridRealOnPosisi And GridX(ShowFixCol(Px2)).GridIndexHead = 0)) Then
        For iX = HscHead.Value To HscHead.Value + (ShowOnConHeadSell - 1)
            If Px1 <= GridX(HeadSellOnFixCol(iX)).GridRealOnPosisi And Px2 >= GridX(HeadSellOnFixCol(iX)).GridRealOnPosisi Then
                If Px1Head = -1 Then Px1Head = iX: Px2Head = iX Else Px2Head = iX
                OpenProsesi = True
                    If Px2 <= GridX(HeadSellOnFixCol(iX)).GridRealOnPosisi Then Exit For
            End If
        Next iX
    End If
End If
'If GridX(Px1).GridRealOnPosisi = -1 Then GridX(Px1).GridRealOnPosisi = Px1
'If GridX(Px2).GridRealOnPosisi = -1 Then GridX(Px2).GridRealOnPosisi = Px2

If GridX(ShowFixCol(Px1)).GridRealOnPosisi = -1 Then GridX(ShowFixCol(Px1)).GridRealOnPosisi = Px1
If GridX(ShowFixCol(Px2)).GridRealOnPosisi = -1 Then GridX(ShowFixCol(Px2)).GridRealOnPosisi = Px2


'MsgBox Px2 > HscHead.Value + ShowOnConHeadSell + Hsc.Value
'MsgBox GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi
''MsgBox GridLeftX Px1 > HscHead.Value + ShowOnConHeadSell
'MsgBox (GridX(Px1).GridIndexHead)
'MsgBox Px1Head & " " & Px2Head
'If GridX(Px1).GridRealOnPosisi >= GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi Then MsgBox ""
'** Posisi Ini Salah = GridX(Index).GridRealOnPosisi
'** Posisi Ini Benar = GridX(HeadSellOnFixCol(Index)).GridRealOnPosisi

'MsgBox Px2
'MsgBox GridX(HeadSellOnFixCol(0)).GridRealOnPosisi
'MsgBox GridX(ShowFixCol(Px1)).GridIndexHead 'Px1 & " " & ShowFixCol(Px1) & " " & GridX(ShowFixCol(Px1)).GridRealOnPosisi
'MsgBox ShowOnConHeadSell + Hsc.Value
'Exit Sub
'GoTo 100
'Shape2.top = SellTop(Py1) + 2
'Shape2.Height = Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3
If ShowOnConHeadSell > 0 Then
''    If (((Px1 >= GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi And GridX(ShowFixCol(Px1)).GridRealOnPosisi <= GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi) _
    Or (Px2 >= GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi And GridX(ShowFixCol(Px2)).GridRealOnPosisi <= GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi)) _
    Or (GridRightX > Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px2 + 1) _
    Or OpenProsesi = True) Then    'And Px1Head > -1 ...... Then
        
    'If Px1 < GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridRealOnPosisi Then MsgBox ""
    
'    If OpenProsesi = True Or _
    ((GridX(ShowFixCol(Px1)).GridIndexHead > 0 And Px1 < GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridRealOnPosisi) Or _
    (GridX(ShowFixCol(Px2)).GridIndexHead > 0 And Px2 < GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridRealOnPosisi)) Or _
    (((GridRightX > Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px2 + 1) And _
    GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1)) Then
        
'    If ((OpenProsesi = True Or _
    (GridX(ShowFixCol(Px1)).GridIndexHead > 0 And Px1 < GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi) Or _
    (GridX(ShowFixCol(Px2)).GridIndexHead > 0 And Px2 < GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi)) Or _
    (GridRightX > Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px2 + 1)) And _
    (GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1) Then
    
    If ((OpenProsesi = True Or _
    (GridX(ShowFixCol(Px1)).GridIndexHead > 0 And Px1 >= GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi And Px1 <= GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi) Or _
    (GridX(ShowFixCol(Px2)).GridIndexHead > 0 And Px2 >= GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi And Px2 <= GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi)) Or _
    (GridRightX > Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px2 + 1)) And _
    (GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1) Then

    
'    If OpenProsesi = True Or _
    (((GridX(ShowFixCol(Px1)).GridIndexHead > 0 And Px1 < GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridRealOnPosisi) Or _
    (GridX(ShowFixCol(Px2)).GridIndexHead > 0 And Px2 < GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridRealOnPosisi)) Or _
    (((GridRightX > Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px2 + 1)) And _
    GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1)) Then
        
        If Px2 <> Px1 Or Py2 <> Py1 Then
            If Px1 > GridX(HeadSellOnFixCol((ShowOnConHeadSell - 1) + HscHead.Value)).GridRealOnPosisi Then
                If Px1 < ShowOnConHeadSell + Hsc.Value Then TsX1n = ShowOnConHeadSell + Hsc.Value Else TsX1n = ShowFixCol(Px1)
            Else
                If GridX(ShowFixCol(Px1)).GridRealOnPosisi < GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi And Px2 <= GridX(HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))).GridRealOnPosisi Then
                    TsX1n = GridX(HeadSellOnFixCol(HscHead.Value)).GridRealOnPosisi
                Else
                    If GridX(ShowFixCol(Px1)).GridIndexHead = 0 And Px1Head > -1 Then TsX1n = HeadSellOnFixCol(Px1Head) Else TsX1n = ShowFixCol(Px1) 'Else
                End If
            End If
            Shape2.left = SellLeft(TsX1n) + 2
            
            If Px2 >= ShowOnConHeadSell + Hsc.Value Then
                If Px2 <= GridRightX Then TsX2n = ShowFixCol(Px2) Else TsX2n = ShowFixCol(GridRightX)
            Else
    '            If Px1Head > -1 Then
                    If GridX(ShowFixCol(Px1)).GridIndexHead = 0 Then TsX1n = HeadSellOnFixCol(Px1Head) Else TsX1n = ShowFixCol(Px1)
                    If GridX(ShowFixCol(Px2)).GridIndexHead = 0 Then TsX2n = HeadSellOnFixCol(Px2Head) Else TsX2n = ShowFixCol(Px2)
                    
                    If GridX(TsX2n).GridIndexHead > ShowOnConHeadSell + HscHead.Value Then
                        TsX2n = HeadSellOnFixCol(HscHead.Value + (ShowOnConHeadSell - 1))
                    End If
                    
                    If SellLeft(TsX2n) + SellWidth(TsX2n) > RangeEndHead And RangeEndHead > 0 Then
                        Wids = SellLeft(TsX2n) + SellWidth(TsX2n) - RangeEndHead + 2
                    End If
    '            End If
            End If
            Shape2.Width = Abs(SellLeft(TsX1n) - SellLeft(TsX2n)) + SellWidth(TsX2n) - Wids
                
            If Py1 >= Vsc.Value Then TsY1n = ShowFixCol(Py1) Else TsY1n = ShowFixCol(Vsc.Value)
            Shape2.top = SellTop(TsY1n) + 2
            
            If Py2 <= GridDownY Then TsY2n = ShowFixCol(Py2) Else TsY2n = ShowFixCol(GridDownY)
            Shape2.Height = Abs(SellTop(TsY1n) - SellTop(TsY2n)) + SellHeight(TsY2n) - 3
            
            If Shape2.Visible = False Then Shape2.Visible = True
        Else
            If Shape2.Visible = True Then Shape2.Visible = False
        End If
    Else
        If Shape2.Visible = True Then Shape2.Visible = False
    End If
Else
    If (GridRightX > Px1 - 1 And Hsc.Value < Px2 + 1) And (GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1) Then
        If Px2 <> Px1 Or Py2 <> Py1 Then
            If Px1 >= Hsc.Value Then TsX1n = ShowFixCol(Px1) Else TsX1n = ShowFixCol(Hsc.Value)
            Shape2.left = SellLeft(TsX1n) + 2
            
            If Px2 <= GridRightX Then TsX2n = ShowFixCol(Px2) Else TsX2n = ShowFixCol(GridRightX)
            Shape2.Width = Abs(SellLeft(TsX1n) - SellLeft(TsX2n)) + SellWidth(TsX2n) - Wids
            
            If Py1 >= Vsc.Value Then TsY1n = ShowFixCol(Py1) Else TsY1n = ShowFixCol(Vsc.Value)
            Shape2.top = SellTop(TsY1n) + 2
            
            If Py2 <= GridDownY Then TsY2n = ShowFixCol(Py2) Else TsY2n = ShowFixCol(GridDownY)
            Shape2.Height = Abs(SellTop(TsY1n) - SellTop(TsY2n)) + SellHeight(TsY2n) - 3
            
            If Shape2.Visible = False Then Shape2.Visible = True
        Else
            If Shape2.Visible = True Then Shape2.Visible = False
        End If
    Else
        If Shape2.Visible = True Then Shape2.Visible = False
    End If
End If

Pxy = iPxy

'Pxy.Px1 = Px1
'Pxy.Py1 = Py1
'Pxy.Px2 = Px2
'Pxy.Py2 = Py2
Exit Sub
100
If OpenProsesi = True Or _
((GridX(ShowFixCol(Px1)).GridIndexHead > 0 And Px1 > HscHead.Value + ShowOnConHeadSell) Or _
(GridX(ShowFixCol(Px2)).GridIndexHead > 0 And Px2 > HscHead.Value + ShowOnConHeadSell)) Or _
(((GridRightX > Px1 - 1 And Hsc.Value + ShowOnConHeadSell < Px2 + 1) And _
GridDownY > Py1 - 1 And Vsc.Value < Py2 + 1)) Then

Command25.Caption = ShowOnConHeadSell

iPxy.Px1 = Px1: iPxy.Py1 = Py1: iPxy.Px2 = Px2: iPxy.Py2 = Py2

    Command30.Caption = ""
    If SellLeft(ShowFixCol(Px1)) = 0 Then
    Px1 = GridLeftX
    Command30.Caption = "OK"
    End If
    If SellTop(Py1) = 0 Then Py1 = GridUpY

    If GridUpY > Py1 Then Py1 = GridUpY
    If Py2 > GridDownY Then Py2 = GridDownY
    
    Cx = SellLeft(ShowFixCol(Px2)) + SellWidth(ShowFixCol(Px2)) - 1
    Cy = SellTop(Py2) + SellHeight(Py2) - 1
        If Cx = 0 Then Cx = SellLeft(ShowFixCol(GridRightX)) + SellWidth(ShowFixCol(GridRightX))
        If Cy = 0 Then Cy = SellTop(GridDownY) + SellHeight(GridDownY)
    
    '|-> Jika Terjadi Pergeseran Pada HSC Atau VSC Secara Back -------------------------------------------------------------------------
    If Hsc.Value > Px1 And GridX(ShowFixCol(Px1)).GridIndexHead = 0 Then GridX(ShowFixCol(Px1)).GridLeft = SetNewGrid.GridSize.GDRangeX
    If Px2 > GridRightX And GridX(ShowFixCol(Px2)).GridIndexHead = 0 Then GridX(ShowFixCol(Px2)).GridLeft = Picture1.ScaleWidth
    If Vsc.Value > Py1 Then GridY(Py1).GridTop = SetNewGrid.GridSize.GDRangeY
    If Py2 > GridDownY Then GridY(Px2).GridHeight = Picture1.ScaleHeight
    '-----------------------------------------------------------------------------------------------------------------------------------
    
    If ShowOnConHeadSell > 0 And Hsc.Value >= Px1 And GridX(ShowFixCol(Px1)).GridIndexHead = 0 Then
        GridX(ShowFixCol(Px1)).GridLeft = GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth + 4
    Else
'        GridX(ShowFixCol(Px1)).GridLeft = GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridLeft + GridX(HeadSellOnFixCol(ShowOnConHeadSell - 1)).GridWidth + 4
    End If
    
    Command25.Caption = Px1 & " " & ShowFixCol(Px1) & " " & GridX(Px1).GridRealOnPosisi
    Shape1.left = SellLeft(ShowFixCol(Px1)) + 2
    Shape1.top = SellTop(Py1) + 2
    If SellWidth(ShowFixCol(Px1)) <= 0 Then SellWidth(ShowFixCol(Px1)) = 3
    Shape1.Width = SellWidth(ShowFixCol(Px1)) - 3 'Abs(SellLeft(Px1) - SellLeft(Px2)) + SellWidth(Px2) - 3
    If SellHeight(Py1) <= 0 Then SellHeight(Py1) = 3
    Shape1.Height = SellHeight(Py1) - 3 'Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3 '25 - 3 'SellHeight(Py2) - 3
    Shape1.Visible = True
    
    If Px2 <> Px1 Or Py2 <> Py1 Then
        'If SellLeft(ShowFixCol(Px1)) <> 25 Then
'        Shape2.left = SellLeft(ShowFixCol(Px1)) + 2
        
If 8 = 9 Then
'        MsgBox Px1Head & " " & Px2Head

'        If ShowOnConHeadSell = 0 And Px1 >= GridLeftX + ShowOnConHeadSell Then
        If ShowOnConHeadSell = 0 Then
            TsX1n = ShowFixCol(Px1)
        Else
'            TsX1n = ShowFixCol(Px1)
'            TsX1n = HeadSellOnFixCol(Px1Head)
            If GridX(Px1).GridIndexHead = 0 Then TsX1n = HeadSellOnFixCol(Px1Head) Else TsX1n = ShowFixCol(Px1) 'Else
        End If
        Shape2.left = SellLeft(TsX1n) + 2
        
        If ShowOnConHeadSell = 0 And Px1 >= GridLeftX + ShowOnConHeadSell Then
            TsX1n = ShowFixCol(Px1)
            TsX2n = ShowFixCol(Px2)
        Else
            If Px1 < ShowOnConHeadSell + Hsc.Value And Px2 < ShowOnConHeadSell + Hsc.Value Then
'                TsX1n = HeadSellOnFixCol(Px1Head)
'                TsX1n = ShowFixCol(Px1)
                If GridX(Px1).GridIndexHead = 0 Then TsX1n = HeadSellOnFixCol(Px1Head) Else TsX1n = ShowFixCol(Px1) 'Else
                
                If GridX(Px2).GridIndexHead = 0 Then TsX2n = HeadSellOnFixCol(Px2Head) Else TsX2n = ShowFixCol(Px2) 'Else
                If SellLeft(HeadSellOnFixCol(Px2Head)) + SellWidth(HeadSellOnFixCol(Px2Head)) > RangeEndHead And RangeEndHead > 0 Then
                    Wids = SellLeft(HeadSellOnFixCol(Px2Head)) + SellWidth(HeadSellOnFixCol(Px2Head)) - RangeEndHead
                End If
            Else
            
            End If
        End If
        
        Shape2.Width = Abs(SellLeft(TsX1n) - SellLeft(TsX2n)) + SellWidth(TsX2n) - Wids
        Shape2.top = SellTop(Py1) + 2
        Shape2.Height = Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3
    If Shape2.Visible = False Then Shape2.Visible = True
End If
        
If 8 = 8 Then
        If ShowOnConHeadSell > 0 And OpenProsesi = True Then
            Shape2.left = SellLeft(HeadSellOnFixCol(Px1Head))
            If Px2 > Hsc.Value + Px2Head Then
                Shape2.Width = Abs(SellLeft(HeadSellOnFixCol(Px1Head)) - SellLeft(ShowFixCol(Px2))) + SellWidth(ShowFixCol(Px2)) '- 25 'SellLeft(ShowFixCol(Px2)) '+ SellWidth(ShowFixCol(Px2))
            Else
                If SellLeft(HeadSellOnFixCol(Px2Head)) + SellWidth(HeadSellOnFixCol(Px2Head)) > RangeEndHead And RangeEndHead > 0 Then
                    Wids = SellLeft(HeadSellOnFixCol(Px2Head)) + SellWidth(HeadSellOnFixCol(Px2Head)) - RangeEndHead
                End If
                Shape2.Width = Abs(SellLeft(HeadSellOnFixCol(Px1Head)) - SellLeft(HeadSellOnFixCol(Px2Head))) + SellWidth(HeadSellOnFixCol(Px2Head)) - Wids
            End If
        Else
'            If ShowOnConHeadSell > 0 And Px1 < Hsc.Value + ShowOnConHeadSell And GridX(ShowFixCol(Px1)).GridIndexHead = 0 Then
            If ShowOnConHeadSell > 0 And Px1 > Hsc.Value + ShowOnConHeadSell Then
'                Shape2.left = SellLeft(HeadSellOnFixCol(ShowOnConHeadSell - 1)) + SellWidth(HeadSellOnFixCol(ShowOnConHeadSell - 1)) '+ 3 + 3
'                Shape2.left = SellLeft(ShowFixCol(ShowOnConHeadSell + GridLeftX)) '+ SellWidth(ShowFixCol(ShowOnConHeadSell + GridLeftX)) '+ 3 + 3
                'Shape2.left = SellLeft(10) + 3
'                MsgBox ShowOnConHeadSell + GridLeftX
            End If
            
            If ShowOnConHeadSell > 0 Then
                If ShowFixCol(Px2) = HeadSellOnFixCol(ShowOnConHeadSell - 1) And RangeEndHead > 0 Then 'MsgBox ""
                    If SellLeft(ShowFixCol(Px2)) + SellWidth(ShowFixCol(Px2)) > RangeEndHead And GridX(Px2).GridIndexHead > 0 Then
                        Wids = SellLeft(ShowFixCol(Px2)) + SellWidth(ShowFixCol(Px2)) - RangeEndHead - 1
                    End If
                End If
            End If
            If ShowFixCol(Px1) > HeadSellOnFixCol(ShowOnConHeadSell - 1) And ShowFixCol(Px1) < ShowFixCol(ShowOnConHeadSell + GridLeftX) Then
                Shape2.left = SellLeft(ShowFixCol(ShowOnConHeadSell + GridLeftX)) '+ SellWidth(ShowFixCol(ShowOnConHeadSell + GridLeftX)) '+ 3 + 3
                Shape2.Width = Abs(SellLeft(ShowFixCol(ShowOnConHeadSell + GridLeftX)) - SellLeft(ShowFixCol(Px2))) + SellWidth(ShowFixCol(Px2))
            Else
                'If Px2 < ShowOnConHeadSell + GridLeftX Then
                Shape2.Width = Abs(SellLeft(ShowFixCol(Px1)) - SellLeft(ShowFixCol(Px2))) + SellWidth(ShowFixCol(Px2)) - 3 - Wids
            End If
        End If
        
        Shape2.Height = Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3 '25 - 3 'SellHeight(Py2) - 3
        If Shape2.Visible = False Then Shape2.Visible = True
End If
    Else
        If Shape2.Visible = True Then Shape2.Visible = False
    End If
Else
    Shape1.Visible = False
    Shape2.Visible = False
End If
'Pusing -Pusing
'GridX(Px1).GridIndexHead = GridX(Px1).GridIndexHead
If ShowOnConHeadSell > 0 And GridX(Px1).GridIndexHead - 1 > 0 And GridX(Px1).GridIndexHead - 1 < ShowOnConHeadSell Then
'    Px1 = -GridX(ShowFixCol(Px1)).GridIndexHead
End If
If ShowOnConHeadSell > 0 And GridX(Px2).GridIndexHead - 1 > 0 And GridX(Px2).GridIndexHead < ShowOnConHeadSell Then
'    Px2 = -GridX(ShowFixCol(Px2)).GridIndexHead
End If

Pxy.Px1 = Px1
Pxy.Py1 = Py1
Pxy.Px2 = Px2
Pxy.Py2 = Py2

'Shape1.Refresh
'Shape2.Refresh
Command21.Caption = Px1Head & " - " & Px2Head
'Command3.Caption = SellLeft(Px1) & " " & Px2 & " " & Cx
'clear oce oye
Exit Sub
Errors:
MsgBox "Error"
End Sub

Private Sub TMPDrawInvertToGrid(ByVal tBGrid As Integer, ByVal ObjMe As Object, iPxy As PointInvert, ByVal Px1 As Long, ByVal Py1 As Long, ByVal Px2 As Long, ByVal Py2 As Long)
Dim TmpPxy As PointInvert
Dim Cx As Long, Cy As Long


If (Px1 > SetNewGrid.GridXCount Or Px2 > SetNewGrid.GridXCount) Or _
   (Py1 > SetNewGrid.GridYCount Or Py2 > SetNewGrid.GridYCount) Then
    Px1 = 0
    Py1 = 0
    Px2 = 0
    Py2 = 0
End If
'Command19.Caption = SetNewGrid.GridXCount & " " & SetNewGrid.GridYCount
'Dim GoOpen As Boolean

' And Hsc.Value < iPxy.Px2 + 1
If ((GridRightX > iPxy.Px1 - 1 And Hsc.Value < iPxy.Px2 + 1) And _
GridDownY > iPxy.Py1 - 1 And Vsc.Value < iPxy.Py2 + 1) Then

iPxy.Px1 = Px1
iPxy.Py1 = Py1
iPxy.Px2 = Px2
iPxy.Py2 = Py2

'    If iPxy.Px2 >= SellCountColumn Then
'        iPxy.Px1 = SellCountColumn - 1
'        iPxy.Px2 = SellCountColumn - 1
'    End If
'    If iPxy.Py2 >= SellCountRow Then
'        iPxy.Py1 = SellCountRow - 1
'        iPxy.Py2 = SellCountRow - 1
'    End If
'GridX(0).GridFront
'MsgBox ShowFixCol(Px1)

    If SellLeft(ShowFixCol(Px1)) = 0 Then Px1 = GridLeftX
    If SellTop(Py1) = 0 Then Py1 = GridUpY

    If GridLeftX > Px1 Then Px1 = GridLeftX
    If Px2 > GridRightX Then Px2 = GridRightX
    
    If GridUpY > Py1 Then Py1 = GridUpY
    If Py2 > GridDownY Then Py2 = GridDownY
    
    Cx = SellLeft(ShowFixCol(Px2)) + SellWidth(ShowFixCol(Px2)) - 1
    Cy = SellTop(Py2) + SellHeight(Py2) - 1
    
    If Cx = 0 Then
        Cx = SellLeft(ShowFixCol(GridRightX)) + SellWidth(ShowFixCol(GridRightX))
    End If
    If Cy = 0 Then
        Cy = SellTop(GridDownY) + SellHeight(GridDownY)
    End If
    
'    Command20.Caption = Px1 & " " & Py1 & " " & Px2 & " " & Py2
'    Command20.Visible = True
    
    
    Shape1.left = SellLeft(ShowFixCol(Px1)) + 2
    Shape1.top = SellTop(Py1) + 2
    If SellWidth(ShowFixCol(Px1)) <= 0 Then SellWidth(ShowFixCol(Px1)) = 3
    Shape1.Width = SellWidth(ShowFixCol(Px1)) - 3 'Abs(SellLeft(Px1) - SellLeft(Px2)) + SellWidth(Px2) - 3
    If SellHeight(Py1) <= 0 Then SellHeight(Py1) = 3
    Shape1.Height = SellHeight(Py1) - 3 'Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3 '25 - 3 'SellHeight(Py2) - 3
    Shape1.Visible = True
    
    If Px2 <> Px1 Or Py2 <> Py1 Then
        Shape2.left = SellLeft(ShowFixCol(Px1)) + 2
        Shape2.top = SellTop(Py1) + 2
        Shape2.Width = Abs(SellLeft(ShowFixCol(Px1)) - SellLeft(ShowFixCol(Px2))) + SellWidth(ShowFixCol(Px2)) - 3
        Shape2.Height = Abs(SellTop(Py1) - SellTop(Py2)) + SellHeight(Py2) - 3 '25 - 3 'SellHeight(Py2) - 3
        Shape2.Visible = True
    Else
        Shape2.Visible = False
    End If
'    Shape1.Refresh
    'Picture1.Cls
'    DrawInvert ObjMe, SellLeft(Px1) + 2, SellTop(Py1) + 2, Cx, Cy

'MsgBox Px2

'    DrawInvert ObjMe, SellLeft(XPointerIndex) + 3, SellTop(YPointerIndex + 2) + 3, _
    Cx - 2, Cy - 2
'End If
Else
    Shape1.Visible = False
    Shape2.Visible = False
End If

Command3.Caption = SellLeft(Px1) & " " & Px2 & " " & Cx
End Sub

Private Sub IndexListS(nXPointerIndex As Long, nYPointerIndex As Long, nIndexList As Integer, nList As Integer)
Dim Test1 As Integer, Test2 As Integer
Dim CountGridList As Integer
Dim Aab As Integer

    XPointerIndex = nXPointerIndex
    YPointerIndex = nYPointerIndex
    IndexList = nIndexList
    Aab = 20
        
    tmpGLShowCount = GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLShowCount - 1
    tmpCountContGList = GridXYData(XPointerIndex, YPointerIndex).GridSubType.CountContGList - 1
    If tmpGLShowCount > -1 Then tmpCountContGList = tmpGLShowCount
    If GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLHeight > 0 Then
        CountList = GridXYData(XPointerIndex, YPointerIndex).GridSubType.GLHeight
    Else
        CountList = ((GridY(YPointerIndex).GridHeight - SellHeight_Def) \ (tmpCountContGList + 1))
    End If

    CountGridList = Int((CountList - Aab) \ GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLRange)
    GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(IndexList).GLPointer = nList

    If nList >= CountGridList Then
'        MsgBox CountGridList & " " & nList
        GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(IndexList).GLScrolIndex = (nList - CountGridList) + 1
    Else
        GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(IndexList).GLScrolIndex = 0
    End If

    Cx = GridX(XPointerIndex).GridLeft + GridX(XPointerIndex).GridWidth - 0
    Cy = GridY(YPointerIndex).GridTop + GridY(YPointerIndex).GridHeight

    If GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList > Cy Then
        YyY = GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList - Cy
    Else
        YyY = 0
    End If

    'GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(IndexList).GLScrolIndex = 1 'GridXYData(XPointerIndex, YPointerIndex).GridSubType.ContGList(UIndexs + IndexList).GLScrolIndex - 1
    
    Picture1.Line (GridX(XPointerIndex).GridLeft + 2, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + 2)- _
    (Cx - 2, GridY(YPointerIndex).GridTop + SellHeight_Def + (CountList * IndexList) + CountList - YyY), vbWhite - 0, BF

    Test1 = XPointerIndex: Test2 = YPointerIndex
    DrawList Test1, Test2, Picture1, IndexList, True
    
End Sub

Sub SaveMe()
'    GetData Picture1.Tag
    Open "D:\DataXXX.txt" For Binary As #1
    Put #1, , SetNewGrid
    Put #1, , GridX()
    Put #1, , GridY()
    Put #1, , GridXYData().GPasswordChar
    Put #1, , GridXYData().Grid
    Put #1, , GridXYData().GridSub
    Put #1, , GridXYData().GridSubType
    Put #1, , GridXYData().GridTag
    Put #1, , GridXYData().GridTagSub
    Put #1, , GridXYData().GridXPicIndex
    Put #1, , GridXYData().GridXYFillColor
    Put #1, , GridXYData().GridXYPicIndex
    Put #1, , GridXYData().GridXYValue
    Put #1, , GridXYData().GridYPicIndex
    Close #1
End Sub

Sub SetInputText_SelLength()
    Text1.SelLength = Len(Text1.Text)
End Sub

Sub SetInputTextSub_SelLength()
    Text2.SelLength = Len(Text2.Text)
End Sub

Sub TestCls()
    Picture1.Cls
End Sub

Sub SetGridToolTipText(GetIt As TPicTollTips, Optional Texts As String, Optional Title As String)
        
    If Texts <> "" Then lblToolTipText.Caption = Texts
    If Title = "" Then lblToolTipText.top = 3 Else lblToolTipText.top = 25
        
    picToolTipText.Height = lblToolTipText.top + lblToolTipText.Height + chkToolTipText.Height + 15
    If (lblToolTipText.left * 2) + lblToolTipText.Width < 150 Then
        picToolTipText.Width = 150
    Else
        picToolTipText.Width = (lblToolTipText.left * 2) + lblToolTipText.Width
    End If
    chkToolTipText.top = picToolTipText.ScaleHeight - chkToolTipText.Height - 3
    
    
    SetWindowRgn picToolTipText.hWnd, CreateRoundRectRgn(0, 0, picToolTipText.ScaleWidth, picToolTipText.Height, 10, 10), True
    shpToolTipText.Width = picToolTipText.Width - 1
    shpToolTipText.Height = picToolTipText.Height - 1
   
    GetIt.nLeft = picToolTipText.left
    GetIt.nWidth = picToolTipText.Width
    GetIt.nTop = picToolTipText.top
    GetIt.nHeight = picToolTipText.Height
    GetIt.nText = lblToolTipText.Caption
    GetIt.nTitle = ""
End Sub

Sub GridShowToolTipText(GetIt As TPicTollTips, Optional Shows As Boolean)

    picToolTipText.left = GetIt.nLeft
    picToolTipText.top = GetIt.nTop
    SetGridToolTipText GetIt
    picToolTipText.Visible = Shows
End Sub


