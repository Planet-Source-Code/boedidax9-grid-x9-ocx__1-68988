Attribute VB_Name = "ORegions"

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|   APIS PARA CREAR EL EFECTO DE PROPORCIONAL PARA EL WALLPAPER                         |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As POINTAPI) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const STRETCH_HALFTONE  As Long = &H4&

Public Type POINTAPI
    X  As Long
    Y  As Long
End Type
Dim picW As Long
Dim picH As Long

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| APIS PARA EFECTO DE CONTORNO DEL FORMULARIO                                           |
'| USADAS PARA TRATAMIENTO DE IMAGENES                                                   |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


Public Sub Read_Config_Skin()
'// procedimento para leer las configuraciones del skin de color y posicion de los
'// botones
On Error Resume Next
  With MusicMp3
     '// Configuracion Mascara Normal
     '---------Titulo de disco y Artista---------------------------
        .lblTrackRuta.ForeColor = Read_INI("Skin", "TitleForeColor", RGB(255, 255, 255), True)
        .lblTrackRuta.left = Read_INI("Skin", "TitleX", 9)
        .lblTrackRuta.top = Read_INI("Skin", "TitleY", 32)
     '--------- Albums ---------------------------------------------
        .picAlbum(1).top = Read_INI("Skin", "AlbumsY", 51)
        .picAlbum(1).left = Read_INI("Skin", "AlbumsX", 9)
    '----------Volumen------------------------------------------------
        .picSliderVol.left = Read_INI("Skin", "VolumeX", 165)
        .picSliderVol.top = Read_INI("Skin", "VolumeY", 62)
    '----------intro---------------------------------------------------
        .imgNormal(5).left = Read_INI("Skin", "IntroX", 35)
        .imgNormal(5).top = Read_INI("Skin", "IntroY", 103)
    '----------silencio------------------------------------------------
        .imgNormal(6).left = Read_INI("Skin", "MuteX", 64)
        .imgNormal(6).top = Read_INI("Skin", "MuteY", 103)
    '----------Repetir-------------------------------------------------
        .imgNormal(7).left = Read_INI("Skin", "RepeatX", 93)
        .imgNormal(7).top = Read_INI("Skin", "RepeatY", 103)
    '----------Aleatorio----------------------------------------------
        .imgNormal(8).left = Read_INI("Skin", "RandomizeX", 121)
        .imgNormal(8).top = Read_INI("Skin", "RandomizeY", 103)
    '----------Bit Rate  -----------------------
        .lblBitrate.ForeColor = Read_INI("Skin", "BitRateForeColor", RGB(255, 255, 255), True)
        .lblBitrate.left = Read_INI("Skin", "BitRateX", 8)
        .lblBitrate.top = Read_INI("Skin", "BitRateY", 118)
    '----------Frequencia  -----------------------
        .lblFreq.ForeColor = Read_INI("Skin", "FreqForeColor", RGB(255, 255, 255), True)
        .lblFreq.left = Read_INI("Skin", "FreqX", 105)
        .lblFreq.top = Read_INI("Skin", "FreqY", 118)
    '----------Rola Actual-----------------------
        .picScroll.ForeColor = Read_INI("Skin", "RolaForeColor", RGB(255, 255, 255), True)
        .picScroll.left = Read_INI("Skin", "RolaX", 10)
        .picScroll.top = Read_INI("Skin", "RolaY", 132)
   '----------Tiempo trascurrido -------------------------------------
        .lblTiempoT.ForeColor = Read_INI("Skin", "TimeActForeColor", RGB(255, 255, 255), True)
        .lblTiempoT.left = Read_INI("Skin", "TimeActX", 9)
        .lblTiempoT.top = Read_INI("Skin", "TimeActY", 148)
   '----------Numero de Tracks en la lista---------------------------
        .lblTrackRep.ForeColor = Read_INI("Skin", "TracksForeColor", RGB(255, 255, 255), True)
        .lblTrackRep.left = Read_INI("Skin", "TracksX", 45)
        .lblTrackRep.top = Read_INI("Skin", "TracksY", 148)
   '----------Tiemp total -------------------------------------------
        .lblDuracion.ForeColor = Read_INI("Skin", "TimeTForeColor", RGB(255, 255, 255), True)
        .lblDuracion.left = Read_INI("Skin", "TimeTX", 134)
        .lblDuracion.top = Read_INI("Skin", "TimeTY", 148)
   '---------- Slider de reproduccion--------------------------------
        .picSliderRep.left = Read_INI("Skin", "PlayerX", 17)
        .picSliderRep.top = Read_INI("Skin", "PlayerY", 161)
   '---------- Anterior track ---------------------------------------
        .imgNormal(0).left = Read_INI("Skin", "PreviousX", 30)
        .imgNormal(0).top = Read_INI("Skin", "PreviousY", 173)
   '----------Play button ------------------------------------------
        .imgNormal(1).left = Read_INI("Skin", "PlayX", 53)
        .imgNormal(1).top = Read_INI("Skin", "PlayY", 173)
   '----------Pause button-------------------------------------------
        .imgNormal(2).left = Read_INI("Skin", "PauseX", 76)
        .imgNormal(2).top = Read_INI("Skin", "PauseY", 173)
   '----------stop button-------------------------------------------
        .imgNormal(3).left = Read_INI("Skin", "StopX", 99)
        .imgNormal(3).top = Read_INI("Skin", "StopY", 173)
   '----------Next button-------------------------------------------
        .imgNormal(4).left = Read_INI("Skin", "NextX", 122)
        .imgNormal(4).top = Read_INI("Skin", "NextY", 173)
   '----------Anterior Album-------------------------------------------
        .imgNormal(9).left = Read_INI("Skin", "PrevAlbumX", 237)
        .imgNormal(9).top = Read_INI("Skin", "PrevAlbumY", 12)
   '----------Cambiar caratula-lista------------------------------------
        .imgNormal(10).left = Read_INI("Skin", "FrontX", 259)
        .imgNormal(10).top = Read_INI("Skin", "FrontY", 12)
   '----------siguiente Album-----------------------------------------
        .imgNormal(11).left = Read_INI("Skin", "NexAlbumX", 282)
        .imgNormal(11).top = Read_INI("Skin", "NexAlbumY", 12)
   '----------minmize button-------------------------------------------
        .imgNormal(12).left = Read_INI("Skin", "MinimizarX", 321)
        .imgNormal(12).top = Read_INI("Skin", "MinimizarY", 12)
   '----------MiniMascara----------------------------------------------
        .imgNormal(13).left = Read_INI("Skin", "MiniMX", 336)
        .imgNormal(13).top = Read_INI("Skin", "MiniMY", 12)
   '----------Close button---------------------------------------------
        .imgNormal(14).left = Read_INI("Skin", "CloseX", 351)
        .imgNormal(14).top = Read_INI("Skin", "CloseY", 12)
   '----------Lsita de Reproduccion y caratula----------------------------
        .ListaRep.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
        .ListaRep.ForeColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
        .ListaRep.left = Read_INI("Skin", "RepX", 176)
        .ListaRep.top = Read_INI("Skin", "RepY", 26)
        .ImagenCaratulA.left = Read_INI("Skin", "RepX", 176)
        .ImagenCaratulA.top = Read_INI("Skin", "RepY", 26)
  '-------------------------------------------------------------------------------
    End With
    '// configuracion Mini mascara---
 With frmMini
  '----------label Tiempo Transcurrido----------------------------
        .lblTiempoT.ForeColor = Read_INI("SkinMini", "TimeActForeColor", RGB(255, 255, 255), True)
        .lblTiempoT.left = Read_INI("SkinMini", "TimeActX", 7)
        .lblTiempoT.top = Read_INI("SkinMini", "TimeActY", 13)
  '---------pic scroll -------------------------------------------
        .picScroll.ForeColor = Read_INI("SkinMini", "RolaForeColor", RGB(255, 255, 255), True)
        .picScroll.left = Read_INI("SkinMini", "RolaX", 47)
        .picScroll.top = Read_INI("SkinMini", "RolaY", 14)
   '---------- Anterior track ---------------------------------------
        .picNormal(0).left = Read_INI("SkinMini", "PreviousX", 183)
        .picNormal(0).top = Read_INI("SkinMini", "PreviousY", 18)
   '----------Play button ------------------------------------------
        .picNormal(1).left = Read_INI("SkinMini", "PlayX", 197)
        .picNormal(1).top = Read_INI("SkinMini", "PlayY", 18)
   '----------Pause button-------------------------------------------
        .picNormal(2).left = Read_INI("SkinMini", "PauseX", 211)
        .picNormal(2).top = Read_INI("SkinMini", "PauseY", 18)
   '----------stop button-------------------------------------------
        .picNormal(3).left = Read_INI("SkinMini", "StopX", 224)
        .picNormal(3).top = Read_INI("SkinMini", "StopY", 18)
   '----------Next button-------------------------------------------
        .picNormal(4).left = Read_INI("SkinMini", "NextX", 237)
        .picNormal(4).top = Read_INI("SkinMini", "NextY", 18)
   '----------Mascara Normal------------------------------------------
        .picNormal(5).left = Read_INI("SkinMini", "NormalMX", 254)
        .picNormal(5).top = Read_INI("SkinMini", "NormalMY", 18)
   '----------Close button---------------------------------------------
        .picNormal(6).left = Read_INI("SkinMini", "CloseX", 266)
        .picNormal(6).top = Read_INI("SkinMini", "CloseY", 18)
 End With
        
        '// si se esta mostrando uno de los siguientes cambiarlos sino pus no
'        If bolDirectoriosShow = True Then
'          frmDirectorios.lstAlbums.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
'          frmDirectorios.lstAlbums.ForeColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
'        End If
        If bolOpcionesShow = True Then
          frmOpciones.ListaSkins.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
          frmOpciones.ListaSkins.ForeColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
        End If
        
        If bolTagsShow = True Then
          frmTags.FileTags.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
          frmTags.FileTags.ForeColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
        End If

        
        If bolLyricsShow = True Then
          frmLyrics.picLyrics.BackColor = Read_INI("Skin", "RepBackColor", RGB(0, 0, 0), True)
          frmLyrics.picBody.BackColor = frmLyrics.picLyrics.BackColor
          frmLyrics.shpFocus.BorderColor = Read_INI("Skin", "RepForeColor", RGB(255, 255, 255), True)
          frmLyrics.lblNoLyrics.ForeColor = frmLyrics.shpFocus.BorderColor
          
          frmLyrics.Order_lblLyrics
          MusicMp3.LyricsIndex = 1
        End If

End Sub

Public Function MakeRegion(picSkin As PictureBox) As Long
 '// procedimento usado para hacer los bordes irregulares del formulario
 '// basado en un picture recorriendo pixel por pixel para buscar las areas
 '// que seran trasparentes o ireegulares
    Dim X As Long, Y As Long, StartLineX As Long
    Dim LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hdc As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hdc = picSkin.hdc
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    '// Leer cual sera el color trasparente para el formulario
     TransparentColor = RGB(255, 0, 255) 'Read_INI("Skin", "ColorTrans", RGB(255, 0, 255), True)
    
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            '// si el pixel es del color trasparente
            If GetPixel(hdc, X, Y) = TransparentColor Or X = PicWidth Then
                '// buscar los pixiles trasparentes
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        '// siempre borrar
                        DeleteObject LineRegion
                    End If
                End If
            Else
                '// buscar los pixeles de no transparente color
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
     MakeRegion = FullRegion
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Sub Combinar_Imagen(Origen As PictureBox, Destino As PictureBox)
  Dim mTranspColor As Long
  mTranspColor = Read_INI("Skin", "ColorTrans", RGB(255, 0, 255), True)
 '// recorrer la picture buscando el color trasparente
 For X = 0 To Origen.ScaleWidth - 1
   For Y = 0 To Origen.ScaleHeight - 1
     If Destino.Point(X, Y) = mTranspColor Then
       Color1 = GetPixel(Origen.hdc, X, Y)
       R = Color1 Mod 256
       b = Int(Color1 / 65536)
       G = (Color1 - (b * 65536) - R) / 256
       SetPixel Destino.hdc, X, Y, RGB(R, G, b)
     End If
   Next Y
   DoEvents
 Next X
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Buttons_Skin()
'// procedimiento para cargar todos los controles, ponerlos en su lugar
'// quitar el color trasparente por defaul rosa (255,0,255) para un mejor efecto
Dim desY As Integer, desAlto As Integer, desAncho As Integer, orgX As Integer, orgAncho As Integer, orgAlto As Integer


With MusicMp3
'----------------------------------------------------------------------------------
  '// ajustar la altura y ancho para los botones de reproduccion
  .picFondo.Width = .picBotones.Width
  .picFondo.Height = .picBotones.Height
 For i = 0 To 4
   .imgNormal(i).Width = .picBotones.Width / 5   '// recorrer boton por boton
   .imgNormal(i).Height = .picBotones.Height / 2 '// el estado normal
   desAncho = .imgNormal(i).ScaleWidth
   desAlto = .imgNormal(i).ScaleHeight
   desX = (i) * (.picFondo.ScaleWidth / 5)
   orgX = .imgNormal(i).left
   orgY = .imgNormal(i).top
   orgAncho = .imgNormal(i).ScaleWidth
   orgAlto = .imgNormal(i).ScaleHeight
   
   desY = 0
   .picFondo.PaintPicture .PicMusic.Image, desX, desY, desAncho, desAlto, orgX, orgY, orgAncho, orgAlto
   desY = .picFondo.ScaleHeight / 2
   .picFondo.PaintPicture .PicMusic.Image, desX, desY, desAncho, desAlto, orgX, orgY, orgAncho, orgAlto
 Next i
 Combinar_Imagen .picFondo, .picBotones
  
'----------------------------------------------------------------------------------
  '// todos los botones de los menus
  .picFondo.Width = .picMenu.Width
  .picFondo.Height = .picMenu.Height
    
 For i = 5 To 14
   .imgNormal(i).Width = .picMenu.Width / 10
   .imgNormal(i).Height = .picMenu.Height / 2
   desAncho = .imgNormal(i).ScaleWidth
   desAlto = .imgNormal(i).ScaleHeight
   desX = (i - 5) * (.picMenu.ScaleWidth / 10)
   orgX = .imgNormal(i).left
   orgY = .imgNormal(i).top
   orgAncho = .imgNormal(i).ScaleWidth
   orgAlto = .imgNormal(i).ScaleHeight
   desY = 0
   .picFondo.PaintPicture .PicMusic.Image, desX, desY, desAncho, desAlto, orgX, orgY, orgAncho, orgAlto
   desY = .picFondo.ScaleHeight / 2
   .picFondo.PaintPicture .PicMusic.Image, desX, desY, desAncho, desAlto, orgX, orgY, orgAncho, orgAlto
 Next i
 Combinar_Imagen .picFondo, .picMenu
  
'----------------------------------------------------------------------------------
  '// los sliders de reproduccion y volumen
  .picFondo.Width = .picSliderRep.Width
  .picFondo.Height = .picSliderRep.Height
  
   desAncho = .picSliderRep.ScaleWidth
   desAlto = .picSliderRep.ScaleHeight
   desX = 0
   orgX = .picSliderRep.left
   orgY = .picSliderRep.top
   orgAncho = .picSliderRep.ScaleWidth
   orgAlto = .picSliderRep.ScaleHeight
   desY = 0
   .picFondo.PaintPicture .PicMusic.Image, desX, desY, desAncho, desAlto, orgX, orgY, orgAncho, orgAlto
   Combinar_Imagen .picFondo, .picSliderRep
      
  .picFondo.Width = .picSliderVol.Width
  .picFondo.Height = .picSliderVol.Height
  
   desAncho = .picSliderVol.ScaleWidth
   desAlto = .picSliderVol.ScaleHeight
   desX = 0
   orgX = .picSliderVol.left
   orgY = .picSliderVol.top
   orgAncho = .picSliderVol.ScaleWidth
   orgAlto = .picSliderVol.ScaleHeight
   desY = 0
   .picFondo.PaintPicture .PicMusic.Image, desX, desY, desAncho, desAlto, orgX, orgY, orgAncho, orgAlto
   Combinar_Imagen .picFondo, .picSliderVol
 '----------------------------------------------------------------------------------
   '// borrar los pictures temporales
   .picFondo.Cls
   .picFondo.Picture = LoadPicture()
 End With
 
'//----------------------------------------------------------------------------------
'// configuracion Pra ajustar los botones de la minimascara -------------------------
With frmMini
 '----------------------------------------------------------------------------------
  '// ajustar la altura y ancho para los botones de reproduccion
  .picFondo.Width = .picBotones.Width
  .picFondo.Height = .picBotones.Height
 For i = 0 To 6
   .picNormal(i).Width = .picBotones.Width / 7   '// recorrer boton por boton
   .picNormal(i).Height = .picBotones.Height / 2 '// el estado normal
   desAncho = .picNormal(i).ScaleWidth
   desAlto = .picNormal(i).ScaleHeight
   desX = (i) * (.picFondo.ScaleWidth / 7)
   orgX = .picNormal(i).left
   orgY = .picNormal(i).top
   orgAncho = .picNormal(i).ScaleWidth
   orgAlto = .picNormal(i).ScaleHeight
   desY = 0
   .picFondo.PaintPicture .picMini.Image, desX, desY, desAncho, desAlto, orgX, orgY, orgAncho, orgAlto
   desY = .picFondo.ScaleHeight / 2
   .picFondo.PaintPicture .picMini.Image, desX, desY, desAncho, desAlto, orgX, orgY, orgAncho, orgAlto
 Next i
 Combinar_Imagen .picFondo, .picBotones
 
    '// borrar los pictures temporales
   .picFondo.Cls
   .picFondo.Picture = LoadPicture()

 End With
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Change_Skin(SkinName As String)
 Dim GraphicsHeight As Integer, desAlto As Integer, desAncho As Integer, orgX As Integer, orgAncho As Integer, orgAlto As Integer
 Dim MiRuta As String, i As Integer
 Dim CursorButton As String
 Dim ExistecurButtons As Boolean
 On Error Resume Next
 
'// verifikar si es el defaul para poner una ruta no valida para que ponga el default
If LCase(SkinName) = "default" Then strSkinSeleccionado = "\Default"
 
MiRuta = Path_Exe(PathSkin)

'----------------------------------------------------------------------------------
'// verificar si existe cada uno de los archivos necesarios para el skin
'// sino poner los de default
'// imagen principal

With MusicMp3
  .PicMusic.Cls
 If Dir(MiRuta & "main.bmp") <> "" Then
  .PicMusic.Picture = LoadPicture(MiRuta & "main.bmp")
 ElseIf Dir(MiRuta & "main.jpg") <> "" Then
      .PicMusic.Picture = LoadPicture(MiRuta & "main.jpg")
    Else
      .PicMusic.Picture = frmPopUp.PicMusic.Picture
    End If
    
'// Imagen para los botones de reproduccion
  .picBotones.Cls
 If Dir(MiRuta & "Buttons.bmp") <> "" Then
  .picBotones.Picture = LoadPicture(MiRuta & "Buttons.bmp")
 ElseIf Dir(MiRuta & "Buttons.jpg") <> "" Then
      .picBotones.Picture = LoadPicture(MiRuta & "Buttons.jpg")
    Else
      .picBotones.Picture = frmPopUp.picBotones.Picture
    End If
     .picBotones.AutoSize = True
'// Imagen para los sliders de reproduccion y volumen
'// Imagen de los albums normal y seleccionado
   .picDiscos.Cls
 If Dir(MiRuta & "Sliders.bmp") <> "" Then
  .picDiscos.Picture = LoadPicture(MiRuta & "Sliders.bmp")
 ElseIf Dir(MiRuta & "Sliders.jpg") <> "" Then
       .picDiscos.Picture = LoadPicture(MiRuta & "Sliders.jpg")
    Else
       .picDiscos.Picture = frmPopUp.picDiscos.Picture
    End If
'// Imagenes para los todos los elementos de menu
   .picMenu.Cls
 If Dir(MiRuta & "Menu.bmp") <> "" Then
  .picMenu.Picture = LoadPicture(MiRuta & "Menu.bmp")
 ElseIf Dir(MiRuta & "Menu.jpg") <> "" Then
       .picMenu.Picture = LoadPicture(MiRuta & "Menu.jpg")
    Else
       .picMenu.Picture = frmPopUp.picMenu.Picture
    End If
    .picMenu.AutoSize = True
'// Imagen de Slider de Reproduccion
  .picSliderRep.Cls
 If Dir(MiRuta & "SliderRep.bmp") <> "" Then
  .picSliderRep.Picture = LoadPicture(MiRuta & "SliderRep.bmp")
 ElseIf Dir(MiRuta & "SliderRep.jpg") <> "" Then
       .picSliderRep.Picture = LoadPicture(MiRuta & "SliderRep.jpg")
    Else
       .picSliderRep.Picture = frmPopUp.picRep.Picture
    End If
'// Imagen de Slider del Volumen
  .picSliderVol.Cls
 If Dir(MiRuta & "SliderVol.bmp") <> "" Then
  .picSliderVol.Picture = LoadPicture(MiRuta & "SliderVol.bmp")
 ElseIf Dir(MiRuta & "SliderVol.jpg") <> "" Then
       .picSliderVol.Picture = LoadPicture(MiRuta & "SliderVol.jpg")
    Else
       .picSliderVol.Picture = frmPopUp.picVol.Picture
    End If
    
'// Imagen de Mini Maskara
   frmMini.picMini.Cls
 If Dir(MiRuta & "Mini.bmp") <> "" Then
   frmMini.picMini.Picture = LoadPicture(MiRuta & "Mini.bmp")
 ElseIf Dir(MiRuta & "Mini.jpg") <> "" Then
       frmMini.picMini.Picture = LoadPicture(MiRuta & "Mini.jpg")
    Else
       frmMini.picMini.Picture = frmPopUp.picMini.Picture
    End If
    
 '// Imagen para los botones de la minimascara
   frmMini.picBotones.Cls
 If Dir(MiRuta & "ButtonsMini.bmp") <> "" Then
   frmMini.picBotones.Picture = LoadPicture(MiRuta & "ButtonsMini.bmp")
 ElseIf Dir(MiRuta & "ButtonsMini.jpg") <> "" Then
       frmMini.picBotones.Picture = LoadPicture(MiRuta & "ButtonsMini.jpg")
    Else
       frmMini.picBotones.Picture = frmPopUp.picBotonesMini.Picture
    End If
   frmMini.picBotones.AutoSize = True
'---------------------------------------------------------------------------------------
'// leer la configuracion del skin de las posiciones de los botones
     Read_Config_Skin
'---------------------------------------------------------------------------------------

'// verifikar si existen los cursores usados para los skins
'// cursor para todos los botones
 If Dir(MiRuta & "curButtons.cur") <> "" Then
   ExistecurButtons = True
   cursorButtons = MiRuta & "curButtons.cur"
 Else
   ExistecurButtons = False
 End If
 
 '// cursor principal
 If Dir(MiRuta & "curMain.cur") <> "" Then
    .PicMusic.MouseIcon = LoadPicture(MiRuta & "curMain.cur")
    frmMini.picMini.MouseIcon = LoadPicture(MiRuta & "curMain.cur")
 Else
   .PicMusic.MouseIcon = LoadPicture()
   frmMini.picMini.MouseIcon = LoadPicture()
 End If

'----------------------------------------------------------------------------------------
' colocar los botones si tienen partes que pueden ser transparentes
   Load_Buttons_Skin
'----------------------------------------------------------------------------------------

 For i = 0 To 16
   If i < 5 Then '// Botones de Reproduccion
      desAncho = .picBotones.ScaleWidth / 5
      desAlto = .picBotones.ScaleHeight / 2
      orgX = (i) * (.picBotones.ScaleWidth / 5)
      orgAncho = .picBotones.ScaleWidth / 5
      orgAlto = .picBotones.ScaleHeight / 2
      GraphicsHeight = 0
      
      'ajustar al ancho de la imagen solo los botones de reproduccion
      .imgNormal(i).Width = .picBotones.Width / 5
      .imgNormal(i).Height = .picBotones.Height / 2
        
      '// Si esta reproduciendo
       If i = 1 And .PlayerIsPlaying = "true" Then GraphicsHeight = .picBotones.ScaleHeight / 2
      '// Si esta Pausado
       If i = 2 And .PlayerIsPlaying = "pause" Then GraphicsHeight = .picBotones.ScaleHeight / 2
      '// Si esta Detenido
       If i = 3 And .PlayerIsPlaying = "false" Then GraphicsHeight = .picBotones.ScaleHeight / 2
        
      .imgNormal(i).PaintPicture .picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
      .PicMusic.PaintPicture .imgNormal(i).Image, .imgNormal(i).left, .imgNormal(i).top, .imgNormal(i).ScaleWidth, .imgNormal(i).ScaleHeight, 0, 0, .imgNormal(i).ScaleWidth, .imgNormal(i).ScaleHeight
      .imgNormal(i).Picture = .imgNormal(i).Image
   
   ElseIf i < 15 Then '// todos los demas botones de menus
        
        desAncho = .picMenu.ScaleWidth / 10
        desAlto = .picMenu.ScaleHeight / 2
        orgX = (i - 5) * (.picMenu.ScaleWidth / 10)
        orgAncho = .picMenu.ScaleWidth / 10
        orgAlto = .picMenu.ScaleHeight / 2
        GraphicsHeight = 0
        'ajustar al ancho de la imagen solo los botones de menus
        .imgNormal(i).Width = .picMenu.Width / 10
        .imgNormal(i).Height = .picMenu.Height / 2
        
        '// si esta intro
        If i = 5 And frmPopUp.mnuIntro.Checked = True Then GraphicsHeight = .picMenu.ScaleHeight / 2
        '// si esta mute
        If i = 6 And frmPopUp.mnuSilencio.Checked = True Then GraphicsHeight = .picMenu.ScaleHeight / 2
        '// si esta en repetir
        If i = 7 And frmPopUp.mnuRepetir.Checked = True Then GraphicsHeight = .picMenu.ScaleHeight / 2
        '// si esta en randomize
        If i = 8 And (frmPopUp.mnuAleatorioActAlbum.Checked = True Or frmPopUp.mnuAleatorioTodaColec.Checked = True) Then
            GraphicsHeight = .picMenu.ScaleHeight / 2
        End If
        '// si esta en caratula
        If i = 10 And MusicMp3.bolShowFront = True Then
         GraphicsHeight = .picMenu.ScaleHeight / 2
        End If
        
        .imgNormal(i).PaintPicture .picMenu.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
        .PicMusic.PaintPicture .imgNormal(i).Image, .imgNormal(i).left, .imgNormal(i).top, .imgNormal(i).ScaleWidth, .imgNormal(i).ScaleHeight, 0, 0, .imgNormal(i).ScaleWidth, .imgNormal(i).ScaleHeight
       
       Else '// Slideres pequeños de reproduccion y volumen
        desAncho = .picDiscos.ScaleWidth / 3
        desAlto = .picDiscos.ScaleHeight / 2
        orgX = (i - 15) * (.picDiscos.ScaleWidth / 3)
        orgAncho = .picDiscos.ScaleWidth / 3
        orgAlto = .picDiscos.ScaleHeight / 2
        .imgNormal(i).PaintPicture .picDiscos.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
        .imgNormal(i).Picture = .imgNormal(i).Image
       End If
       
   '// Colokar los cursores si existen a todos los botones
   If ExistecurButtons = True Then
     .imgNormal(i).MouseIcon = LoadPicture(cursorButtons)
   Else
     .imgNormal(i).MouseIcon = LoadPicture()
   End If
     
Next i

For i = 0 To 6   '// Botones de Reproduccion de mini mascara
      desAncho = frmMini.picBotones.ScaleWidth / 7
      desAlto = frmMini.picBotones.ScaleHeight / 2
      orgX = (i) * (frmMini.picBotones.ScaleWidth / 7)
      orgAncho = frmMini.picBotones.ScaleWidth / 7
      orgAlto = frmMini.picBotones.ScaleHeight / 2
      GraphicsHeight = 0
      
      frmMini.picNormal(i).Width = frmMini.picBotones.Width / 7
      frmMini.picNormal(i).Height = frmMini.picBotones.Height / 2
        
      '// Si esta reproduciendo
       If i = 1 And .PlayerIsPlaying = "true" Then GraphicsHeight = frmMini.picBotones.ScaleHeight / 2
      '// Si esta Pausado
       If i = 2 And .PlayerIsPlaying = "pause" Then GraphicsHeight = frmMini.picBotones.ScaleHeight / 2
      '// Si esta Detenido
       If i = 3 And .PlayerIsPlaying = "false" Then GraphicsHeight = frmMini.picBotones.ScaleHeight / 2
        
      '// copiar imagen a los pictures individuales
      frmMini.picNormal(i).PaintPicture frmMini.picBotones.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
      '// copiar imagen de los pictures individuales al fondo
      frmMini.picMini.PaintPicture frmMini.picNormal(i).Image, frmMini.picNormal(i).left, frmMini.picNormal(i).top, frmMini.picNormal(i).ScaleWidth, frmMini.picNormal(i).ScaleHeight, 0, 0, frmMini.picNormal(i).ScaleWidth, frmMini.picNormal(i).ScaleHeight
      frmMini.picNormal(i).Picture = frmMini.picNormal(i).Image
      
         '// Colokar los cursores si existen a todos los botones
      If ExistecurButtons = True Then
        frmMini.picNormal(i).MouseIcon = LoadPicture(cursorButtons)
      Else
        frmMini.picNormal(i).MouseIcon = LoadPicture()
      End If
 Next i
   
   '// imagen del picture del scrolltex
   orgX = .picScroll.left
   orgY = .picScroll.top
   .picScroll.PaintPicture .PicMusic.Image, 0, 0, .picScroll.ScaleWidth, .picScroll.ScaleHeight, orgX, orgY, .picScroll.ScaleWidth, .picScroll.ScaleHeight
   .picScroll.Picture = .picScroll.Image
   
   If ExistecurButtons = True Then
      '// colokar cursores de botones al picscrolltext
      '// y a los botones de la mini mascara
      .picScroll.MouseIcon = LoadPicture(cursorButtons)
      frmMini.picScroll.MouseIcon = LoadPicture(cursorButtons)
   Else
     .picScroll.MouseIcon = LoadPicture()
     frmMini.picScroll.MouseIcon = LoadPicture()
   End If
   
   '// imagen del picture del scrolltext para la mini mascara
   orgX = frmMini.picScroll.left
   orgY = frmMini.picScroll.top
   frmMini.picScroll.PaintPicture frmMini.picMini.Image, 0, 0, frmMini.picScroll.ScaleWidth, frmMini.picScroll.ScaleHeight, orgX, orgY, frmMini.picScroll.ScaleWidth, frmMini.picScroll.ScaleHeight
   frmMini.picScroll.Picture = frmMini.picScroll.Image
'-----------------------------------------------------------------------------------------
 '// Rotar el texto de nuevo con la nueva mascara
  If bolMiniMascara = True Then
    .Scroll_Text ScrollText, frmMini.picScroll
  Else
    .Scroll_Text ScrollText, MusicMp3.picScroll
  End If
'+---------------------------------------------------------------------------------------+

'// cursor para los albums
 If Dir(MiRuta & "curAlbums.cur") <> "" Then
    .picAlbum(1).MouseIcon = LoadPicture(MiRuta & "curAlbums.cur")
 ElseIf ExistecurButtons = True Then
        .picAlbum(1).MouseIcon = LoadPicture(cursorButtons)
     Else
        .picAlbum(1).MouseIcon = LoadPicture()
     End If

'// cargar los albums con el nuevo skin segun la posicion
For i = 1 To TotalAlbumS
  'si es el primer album se queda en la misma posicion
 If i <= 48 Then  ' comparar los albums que se pueden ver maximo 48
  If i <> 1 And i < 13 Then '// primera linea de 12 elementos
    .picAlbum(i).top = .picAlbum(1).top
    .picAlbum(i).left = .picAlbum(i - 1).left + 13
  End If
  
  If i > 12 And i < 25 Then '// Segunda linea de 12 elementos
   .picAlbum(i).top = .picAlbum(1).top + 13
   .picAlbum(i).left = .picAlbum(i - 12).left
  End If
  
  If i > 24 And i < 37 Then '// Tercera linea de 12 elementos
   .picAlbum(i).top = .picAlbum(13).top + 13
   .picAlbum(i).left = .picAlbum(i - 24).left
  End If
  
  If i > 36 And i < 49 Then '// Cuarta linea de 12 elementos
   .picAlbum(i).top = .picAlbum(25).top + 13
   .picAlbum(i).left = .picAlbum(i - 36).left
  End If
  
 '// Poner la imagen ahora si
  desAncho = .picDiscos.ScaleWidth / 3
  desAlto = .picDiscos.ScaleHeight / 2
  orgX = (2) * (.picDiscos.ScaleWidth / 3)
  orgAncho = .picDiscos.ScaleWidth / 3
  orgAlto = .picDiscos.ScaleHeight / 2
  
  If intActiveAlbum = i Then '// activar el album activo
    GraphicsHeight = .picDiscos.ScaleHeight / 2
  Else
    GraphicsHeight = 0
  End If
  
  .picAlbum(i).PaintPicture .picDiscos.Image, 0, 0, desAncho, desAlto, orgX, GraphicsHeight, orgAncho, orgAlto
  .picAlbum(i).MouseIcon = .picAlbum(1).MouseIcon
 End If

Next i
  
 '// cursor de la imagen del slider de reproduccion
 If Dir(MiRuta & "cursliderrep.cur") <> "" Then
    .imgNormal(15).MouseIcon = LoadPicture(MiRuta & "curSliderrep.cur")
 ElseIf Dir(MiRuta & "curslidervol.cur") <> "" Then
       '// si no tiene cursor poner el del volumen
        .imgNormal(15).MouseIcon = LoadPicture(MiRuta & "curSlidervol.cur")
     Else
       '// si no hay ninguno porner el de los menus si tiene
        .imgNormal(15).MouseIcon = .imgNormal(0).MouseIcon
     End If
     
 '// cursor de la imagen del slider de volumen
 If Dir(MiRuta & "curslidervol.cur") <> "" Then
     .imgNormal(16).MouseIcon = LoadPicture(MiRuta & "curSlidervol.cur")
 Else
     .imgNormal(16).MouseIcon = .imgNormal(15).MouseIcon
 End If
     
 '// Cursor para la lista de reproduccion
 If Dir(MiRuta & "curListaRep.cur") <> "" Then
    .ListaRep.MouseIcon = LoadPicture(MiRuta & "curListaRep.cur")
 Else
   .ListaRep.MouseIcon = .imgNormal(15).MouseIcon
 End If
 
   .PicMusic.PaintPicture .picSliderRep.Image, .picSliderRep.left, .picSliderRep.top, .picSliderRep.ScaleWidth, .picSliderRep.ScaleHeight, 0, 0, .picSliderRep.ScaleWidth, .picSliderRep.ScaleHeight
   .PicMusic.PaintPicture .picSliderVol.Image, .picSliderVol.left, .picSliderVol.top, .picSliderVol.ScaleWidth, .picSliderVol.ScaleHeight, 0, 0, .picSliderVol.ScaleWidth, .picSliderVol.ScaleHeight

 End With
 MusicMp3.PicMusic.Refresh
 frmMini.picMini.Refresh

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'// procedimento para hacer calkular la maskara normal y la mini
Sub Form_Mini_Normal(Obj As Object)
    
    'RegionMusic = MakeRegion(frmServerTest.Picture3(6))
    'SetWindowRgn frmServerTest.Picture3(6).HwnD, RegionMusic, True
    
    RegionMusic = MakeRegion(Obj)
    SetWindowRgn Obj.hWnd, RegionMusic, True
 
 
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Skins_Menu(SelMenu As String)
'// Procedimiento para cargar los skins disponibles que son todos las carpetas
'// en la ruta del EXE mas \MMp3Player\Skins\  y los carga en el menu de frmpopup
'// parametros
'// [SelMenu] -> Menu el cual va ha estar seleccionado

Dim miNombre As String
Dim i As Integer
On Error Resume Next
MiRuta = Path_Exe(PathExe) & "MMp3Player\Skins\"
i = 1
miNombre = Dir(MiRuta, vbDirectory)   ' Recupera la primera entrada.

'/* para ver si hay imagnes en el directorio
frmPopUp.fileBmps.Pattern = "*.bmp;*.jpg"

frmPopUp.mnuSkinsAdd(1).Checked = True

Do While miNombre <> ""
   If miNombre <> "." And miNombre <> ".." Then
      ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
      If (GetAttr(MiRuta & miNombre) And vbDirectory) = vbDirectory Then
       frmPopUp.fileBmps.Path = MiRuta & miNombre
        If frmPopUp.fileBmps.ListCount > 0 Then
             i = i + 1
             Load frmPopUp.mnuSkinsAdd(i)  '// cargar los menus dinamikamente
             frmPopUp.mnuSkinsAdd(i).Caption = " " & miNombre
             frmPopUp.mnuSkinsAdd(i).Checked = False
          If LCase(miNombre) = SelMenu Then
            frmPopUp.mnuSkinsAdd(1).Checked = False
            frmPopUp.mnuSkinsAdd(i).Checked = True
          End If
        End If
      End If
   End If
  miNombre = Dir
Loop
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'+--------------------------------------------------------------------------------------+
'|    CREAR LA IMAGEN DE WALLPAPER SEGUN LAS OPCIONES ESPECIFICADAS                     |
'+--------------------------------------------------------------------------------------+

Public Sub CreatePic(picSource As PictureBox, picDestination As PictureBox)
'// Procedimiento para krear el strech con la mas alta calidad posible
Dim hBrush          As Long
Dim hDummyBrush     As Long
Dim lOrigMode       As Long
Dim uBrushOrigPt    As POINTAPI
Dim lWidth As Long
Dim lHeight As Long
Dim lLeft As Integer
Dim lTop As Integer
    picDestination.AutoRedraw = True
    picDestination.Cls
    lWidth = picDestination.Width
    lHeight = picDestination.Height
    lLeft = 0
    lTop = 0
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(picDestination.hdc, STRETCH_HALFTONE)

    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(0)
    hBrush = SelectObject(picDestination.hdc, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    UnrealizeObject hBrush
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    SetBrushOrgEx picDestination.hdc, lLeft, lTop, uBrushOrigPt
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(picDestination.hdc, hBrush)
    
    'Stretch the bitmap
    StretchBlt picDestination.hdc, lLeft, lTop, lWidth, lHeight, _
            picSource.hdc, 0, 0, picSource.Width, picSource.Height, vbSrcCopy
    
    'Set the stretch mode back to it's original mode
    SetStretchBltMode picDestination.hdc, lOrigMode
    
    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(picDestination.hdc, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    UnrealizeObject hBrush
    'Set the brush alignment back to the original coordinates
    SetBrushOrgEx picDestination.hdc, uBrushOrigPt.X, uBrushOrigPt.Y, uBrushOrigPt
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(picDestination.hdc, hBrush)
    'Get rid of the dummy brush
    DeleteObject hDummyBrush
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+


'+--------------------------------------------------------------------------------------+
'|    CREAR LA IMAGEN DE WALLPAPER Y PONER EN EL ESCRITORIO                             |
'+--------------------------------------------------------------------------------------+

Public Sub ConfigurarWallpaper()
'// procedimiento para krear la imagen y ponerla en el escritorio como wallpaper
  On Error GoTo Bitch
    If frmPopUp.mnuWallpapper.Checked = False Then Exit Sub
       MusicMp3.picWallOriginal.Picture = Nothing
       MusicMp3.picWallOriginal.Width = 1
       MusicMp3.picWallOriginal.Height = 1
       
        If OpcionesMusic.NoAlteraR = True Then Exit Sub
         If Trim(strRutaCaratula) = "" Then '// no tiene caratula poner el default
           If bolCaratulaDefault = True Then Exit Sub '// ponerla solo una vez
           MusicMp3.picWallOriginal.Picture = frmCaratula.Picture2.Picture
           SavePicture MusicMp3.picWallOriginal.Image, DirectoriOWindowS & "MusicMp3.bmp"
           PoneRWallPapeR "Mosaico"
           bolCaratulaDefault = True
           GoTo Bitch
         Else  'si tiene caratula ponerla
           MusicMp3.picWallOriginal.Picture = LoadPicture(strRutaCaratula)
           bolCaratulaDefault = False
         End If
          
         '// Wallpaper estilo Expandido
         If OpcionesMusic.Expander Then
           SavePicture MusicMp3.picWallOriginal.Image, DirectoriOWindowS & "MusicMp3.bmp"
           PoneRWallPapeR "Expandido"
           Exit Sub
         End If
         
         '// Wallpaper Stylo proporcional
         If OpcionesMusic.Proporcional = True Then
            '----ajustar la ..che imagen para que quede chida-----------------------
            MusicMp3.picWallProp.Picture = Nothing
             If MusicMp3.picWallOriginal.Width > MusicMp3.picWallOriginal.Height Then
               MusicMp3.picWallProp.Width = Screen.Width
               MusicMp3.picWallProp.Height = MusicMp3.picWallOriginal.Height * Screen.Width / MusicMp3.picWallOriginal.Width
             Else
               MusicMp3.picWallProp.Height = Screen.Height
               MusicMp3.picWallProp.Width = MusicMp3.picWallOriginal.Width * Screen.Height / MusicMp3.picWallOriginal.Height
             End If
               CreatePic MusicMp3.picWallOriginal, MusicMp3.picWallProp
            '----------------------------------------------------------------------
            SavePicture MusicMp3.picWallProp.Image, DirectoriOWindowS & "MusicMp3.bmp"
              '// Wallpaper estilo Centrado
               If OpcionesMusic.Centrar = True Then
                 PoneRWallPapeR "Centro"
                 GoTo Bitch
               End If
              '// Wallpaper Estilo Mosaiko
               If OpcionesMusic.Mosaico = True Then
                 PoneRWallPapeR "Mosaico"
                 GoTo Bitch
               End If
         Else
            '// si no es proporcional
            SavePicture MusicMp3.picWallOriginal.Image, DirectoriOWindowS & "MusicMp3.bmp"
               If OpcionesMusic.Centrar = True Then
                 PoneRWallPapeR "Centro"
                 GoTo Bitch
               End If
               If OpcionesMusic.Mosaico = True Then
                 PoneRWallPapeR "Mosaico"
                 GoTo Bitch
               End If
         End If
Exit Sub
Bitch:
    MusicMp3.picWallOriginal.Picture = Nothing
    MusicMp3.picWallProp.Picture = Nothing
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+








Public Function GetGreen(ByVal ColorVal As Long) As Integer
GetGreen = ((ColorVal And &HFF00FF00) / 256&)
End Function

Public Function GetBlue(ByVal ColorVal As Long) As Integer
GetBlue = (ColorVal And &HFF0000) / (256& * 256&)
End Function

Public Function GetRed(ByVal ColorVal As Long) As Integer
GetRed = ColorVal Mod 256
End Function

Public Function Mix_Color(ParamArray Colors()) As Long
Dim Lb, UB As Integer
Dim Red, Green, Blue
  
Lb = LBound(Colors)
UB = UBound(Colors)

For i = Lb To UB
    Red = Red + GetRed(Colors(i))
    Green = Green + GetGreen(Colors(i))
    Blue = Blue + GetBlue(Colors(i))
Next

Red = Red / (UB + 1)
Blue = Blue / (UB + 1)
Green = Green / (UB + 1)

Mix_Color = RGB(Red, Green, Blue)
End Function

Public Function InvertColors(Colors As Long) As Long
Dim Red, Green, Blue

Red = GetRed(Colors) Xor 250
Green = GetGreen(Colors) Xor 250
Blue = GetBlue(Colors) Xor 250

InvertColors = RGB(Red, Green, Blue)
End Function



