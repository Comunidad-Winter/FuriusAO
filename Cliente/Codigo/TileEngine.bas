Attribute VB_Name = "Mod_TileEngine"

Option Explicit







Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521


Public Const SRCCOPY = &HCC0020









Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type


Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type


Public Type Position
    X As Integer
    Y As Integer
End Type


Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type



Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type


Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type


Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type


Public Type HeadData
    Head(1 To 4) As Grh
End Type


Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
End Type


Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type



Public Type FxData
    FX As Grh
    OffsetX As Long
    OffsetY As Long
End Type


Public Type Char
    Active As Byte
    Heading As Byte
    POS As Position

    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    FX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    Navegando As Byte
    
    Nombre As String
    
    haciendoataque As Byte
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    
End Type


Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type


Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh

    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type


Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    
    Changed As Byte
End Type


Public IniPath As String
Public MapPath As String



Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte


Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long


Public WindowTileWidth As Integer
Public WindowTileHeight As Integer


Public MainViewTop As Integer
Public MainViewLeft As Integer




Public TileBufferSize As Integer


Public DisplayFormhWnd As Long


Public TilePixelHeight As Integer
Public TilePixelWidth As Integer



Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer



Public LastTime As Long



Public MainDestRect   As RECT

Public MainViewRect   As RECT
Public BackBufferRect As RECT




Public MainViewWidth As Integer
Public MainViewHeight As Integer





Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh



Public MapData() As MapBlock
Public MapInfo As MapInfo



Public CharList(1 To 10000) As Char




Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, _
     ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, _
     ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer
Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
     ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
     ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
     ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
     ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
'Code by Ladder 11/12/07




Public bRain        As Boolean
Public bRainST      As Boolean
Public bTecho       As Boolean
'leito nieve
Public bSnowST      As Boolean
Public bSnow        As Boolean 'está nevando?
Public bNoche       As Boolean 'es de noche?

Public brstTick     As Long

' Leito nieve
Private RNieva(7)  As RECT  'RECT de la nieve
Private RLluvia(7)  As RECT
Private iFrameIndex As Byte
Private llTick      As Long
Private LTLluvia(4) As Integer
'Leito nieve
Private LTNieva(4) As Integer


Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum


    Public Enum PlayLoop
        plNone = 0
        plLluviain = 1
        plLluviaout = 2
        plFogata = 3
    End Enum





Sub CargarCabezas()
On Error Resume Next
Dim n As Integer, i As Integer, Numheads As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , Numheads


ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #n, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCascos()
On Error Resume Next
Dim n As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCascos


ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #n, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

n = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCuerpos


ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #n, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
Next i

Close #n

End Sub
Sub CargarFxs()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

n = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumFxs


ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #n, , MisFxs(i)
    Call InitGrh(FxData(i).FX, MisFxs(i).Animacion, 1)
    FxData(i).OffsetX = MisFxs(i).OffsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #n

End Sub
Sub CargarArrayLluvia()
On Error Resume Next
Dim n As Integer, i As Integer
Dim Nu As Integer

n = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #n
'cabecera
Get #n, , MiCabecera
'num de cabezas
Get #n, , Nu
'Resize array
'ReDim bLluvia(1 To Nu) As Byte
'For I = 1 To Nu
'    Get #N, , bLluvia(I)
'Next I
Close #n

ReDim bLluvia(1 To Nu) As Byte
bLluvia(1) = 1
bLluvia(2) = 1
bLluvia(3) = 1
bLluvia(4) = 1
bLluvia(5) = 1
bLluvia(6) = 1
bLluvia(7) = 1
bLluvia(8) = 1
bLluvia(9) = 1
bLluvia(10) = 1
bLluvia(11) = 1
bLluvia(12) = 1
bLluvia(13) = 1
bLluvia(14) = 1
bLluvia(15) = 1
bLluvia(16) = 1
bLluvia(17) = 1
bLluvia(18) = 1
bLluvia(19) = 1
bLluvia(20) = 1
bLluvia(21) = 1
bLluvia(22) = 1
bLluvia(23) = 1
bLluvia(24) = 1
bLluvia(25) = 1
bLluvia(26) = 1
bLluvia(27) = 1
bLluvia(28) = 1
bLluvia(29) = 1
bLluvia(30) = 1
bLluvia(31) = 1
bLluvia(32) = 1
bLluvia(34) = 1
bLluvia(35) = 1
bLluvia(36) = 1
bLluvia(38) = 1
bLluvia(39) = 1
bLluvia(46) = 1
bLluvia(47) = 1
bLluvia(53) = 1
bLluvia(54) = 1
bLluvia(55) = 1
bLluvia(56) = 1
bLluvia(57) = 1
bLluvia(58) = 1
bLluvia(59) = 1
bLluvia(60) = 1
bLluvia(61) = 1
bLluvia(62) = 1
bLluvia(63) = 1
bLluvia(64) = 1
bLluvia(65) = 1
bLluvia(66) = 1
bLluvia(67) = 1
bLluvia(68) = 1
bLluvia(69) = 1
bLluvia(70) = 1
bLluvia(71) = 1
bLluvia(72) = 1
bLluvia(73) = 1
bLluvia(74) = 1
bLluvia(75) = 1
bLluvia(76) = 1
bLluvia(77) = 0
bLluvia(78) = 1
bLluvia(79) = 1
bLluvia(80) = 1
bLluvia(81) = 1
bLluvia(82) = 1
bLluvia(83) = 1
bLluvia(84) = 1
bLluvia(85) = 1
bLluvia(86) = 1
bLluvia(87) = 1
bLluvia(88) = 1
bLluvia(89) = 1
bLluvia(90) = 1
bLluvia(91) = 1
bLluvia(92) = 1
bLluvia(93) = 1
bLluvia(94) = 1
bLluvia(95) = 1
bLluvia(96) = 1
bLluvia(97) = 1
bLluvia(98) = 1
bLluvia(99) = 1
bLluvia(100) = 1
bLluvia(101) = 1
bLluvia(102) = 1
bLluvia(103) = 1
bLluvia(104) = 1
bLluvia(105) = 1
bLluvia(106) = 1
bLluvia(107) = 1
bLluvia(108) = 1
bLluvia(109) = 1
bLluvia(110) = 1
bLluvia(111) = 1
bLluvia(112) = 1
bLluvia(113) = 1
bLluvia(114) = 1
bLluvia(117) = 1
bLluvia(118) = 1
bLluvia(119) = 1
bLluvia(120) = 1
bLluvia(121) = 1
bLluvia(122) = 1
bLluvia(123) = 1
bLluvia(124) = 1
bLluvia(125) = 1
bLluvia(126) = 1
bLluvia(127) = 1
bLluvia(128) = 1
bLluvia(129) = 1
bLluvia(130) = 1
bLluvia(131) = 1
bLluvia(131) = 1
bLluvia(132) = 1
bLluvia(133) = 1
bLluvia(134) = 1
bLluvia(135) = 1
bLluvia(136) = 1
bLluvia(137) = 1
bLluvia(138) = 1
bLluvia(139) = 1
bLluvia(147) = 1
bLluvia(148) = 1
bLluvia(149) = 1
bLluvia(150) = 1
bLluvia(151) = 1
bLluvia(152) = 1
bLluvia(153) = 1
bLluvia(154) = 1
bLluvia(155) = 1
bLluvia(156) = 1
bLluvia(157) = 1
bLluvia(158) = 1
bLluvia(159) = 1
bLluvia(160) = 1
bLluvia(161) = 1
bLluvia(162) = 1
bLluvia(173) = 1
bLluvia(177) = 1

ReDim bNieva(1 To Nu) As Byte

bNieva(170) = 1
bNieva(169) = 1
bNieva(171) = 1

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal cx As Single, ByVal cy As Single, tX As Integer, tY As Integer)




Dim HWindowX As Integer
Dim HWindowY As Integer

cx = cx - StartPixelLeft
cy = cy - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)


cx = (cx \ TilePixelWidth)
cy = (cy \ TilePixelHeight)

If cx > HWindowX Then
    cx = (cx - HWindowX)
Else
    If cx < HWindowX Then
        cx = (0 - (HWindowX - cx))
    Else
        cx = 0
    End If
End If

If cy > HWindowY Then
    cy = cy + 0 - HWindowY
Else
    If cy < HWindowY Then
        cy = (cy - HWindowY)
    Else
        cy = 0
    End If
End If

tX = UserPos.X + cx
tY = UserPos.Y + cy

End Sub
Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer)
On Error Resume Next


If CharIndex > LastChar Then LastChar = CharIndex

NumChars = NumChars + 1

If arma = 0 Then arma = 2
If escudo = 0 Then escudo = 2
If casco = 0 Then casco = 2

CharList(CharIndex).Head = HeadData(Head)

CharList(CharIndex).Body = BodyData(Body)

If Body > 83 And Body < 88 Then
    CharList(CharIndex).Navegando = 1
Else: CharList(CharIndex).Navegando = 0
End If

CharList(CharIndex).arma = WeaponAnimData(arma)
    
CharList(CharIndex).escudo = ShieldAnimData(escudo)
CharList(CharIndex).casco = CascoAnimData(casco)

CharList(CharIndex).Heading = Heading


CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0


CharList(CharIndex).POS.X = X
CharList(CharIndex).POS.Y = Y


CharList(CharIndex).Active = 1


MapData(X, Y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

CharList(CharIndex).Active = 0
CharList(CharIndex).Criminal = 0
CharList(CharIndex).FX = 0
CharList(CharIndex).FxLoopTimes = 0
CharList(CharIndex).invisible = False
CharList(CharIndex).Moving = 0
CharList(CharIndex).muerto = False
CharList(CharIndex).Nombre = ""
CharList(CharIndex).pie = False
CharList(CharIndex).POS.X = 0
CharList(CharIndex).POS.Y = 0
CharList(CharIndex).UsandoArma = False

End Sub

Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next





CharList(CharIndex).Active = 0


If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).CharIndex = 0

Call ResetCharInfo(CharIndex)


NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)



If GrhIndex = 0 Then Exit Sub
Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1







If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed



End Sub

Sub MoveCharByHead(CharIndex As Integer, nheading As Byte)



Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y


Select Case nheading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading

If UserEstado <> 1 Then Call DoPasosFx(CharIndex)

Call frmMmap.MoverPJ

End Sub


Public Sub DoFogataFx()
If FX = 0 Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then frmMain.StopSound
    Else
        bFogata = HayFogata()
        If bFogata Then frmMain.Play "fuego.wav", True
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim X As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next X
Next Y

EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
Static TickCount(200) As Integer
If Cont = 999 Then Exit Function
TickCount(Cual) = TickCount(Cual) + 1
If TickCount(Cual) < Cont Then
    TickON = False
Else
    TickCount(Cual) = 0
    TickON = True
End If
End Function
Sub DoPasosFx(ByVal CharIndex As Integer)
Static pie As Boolean

If CharList(CharIndex).Navegando = 0 Then
   ' If UserMontando And EstaPCarea(CharIndex) And CharIndex = UserCharIndex Then
   '     If TickON(0, 4) Then Call PlayWaveDS(SND_MONTANDO)
   ' Else
        If CharList(CharIndex).Criminal = 1 Or CharList(CharIndex).Criminal > 5 And CharList(CharIndex).Criminal < 9 Then Exit Sub
        If Not CharList(CharIndex).muerto And EstaPCarea(CharIndex) Then
            CharList(CharIndex).pie = Not CharList(CharIndex).pie
            If CharList(CharIndex).pie Then
                Call PlayWaveDS(SND_PASOS1)
            Else
                Call PlayWaveDS(SND_PASOS2)
            End If
        End If
   ' End If
Else: Call PlayWaveDS(SND_NAVEGANDO)
End If

End Sub
Sub MoveCharByPosAndHead(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)

On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer



X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y




MapData(nX, nY).CharIndex = CharIndex


CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading




End Sub
Sub MoveCharByPos(CharIndex As Integer, nX As Integer, nY As Integer)
On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nheading As Byte

X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y


If Sgn(addX) = 1 Then nheading = EAST
If Sgn(addX) = -1 Then nheading = WEST
If Sgn(addY) = -1 Then nheading = NORTH
If Sgn(addY) = 1 Then nheading = SOUTH

MapData(nX, nY).CharIndex = CharIndex

CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading

End Sub
Sub MoveCharByPosConHeading(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
On Error Resume Next

If InMapBounds(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y) Then MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).CharIndex = 0

MapData(nX, nY).CharIndex = CharIndex

CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0

CharList(CharIndex).Heading = nheading

End Sub

Sub MoveScreen(Heading As Byte)



Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer


Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select


tX = UserPos.X + X
tY = UserPos.Y + Y


If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.X - 8 To UserPos.X + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
Sub LoadGrhData()
On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer



ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

Get #1, , Grh

Do Until Grh <= 0
    
    
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If
    
    
    Get #1, , Grh

Loop


Close #1
'CARGAR ANIM EXTRAS
GrhData(12460).FileNum = 1588
GrhData(12481).FileNum = 1587
'/
Exit Sub









ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    
   If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
       Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) And ESGM = False Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function

Function LegalPosMuerto(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPosMuerto = False
    Exit Function
End If

    
    If MapData(X, Y).Blocked = 1 Then
        LegalPosMuerto = False
        Exit Function
    End If
    
    
    If MapData(X, Y).CharIndex > 0 Then
    If CharList(MapData(X, Y).CharIndex).muerto = True Then
        LegalPosMuerto = False
        Exit Function
    End If
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    End If
    
LegalPosMuerto = True

End Function




Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(X As Integer, Y As Integer) As Boolean




If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
surface.BltFast X, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(surface As DirectDrawSurface7, Grh As Integer, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

surface.GetSurfaceDesc SurfaceDesc


If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    surface.BltFast destRect.Left, destRect.Top, SurfaceDB.GetBMP(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub



Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)



Dim iGrhIndex As Integer

Dim SourceRect As RECT

Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub
'Exit Sub
'End If

iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
'surface.BltFast X, Y, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_DESTCOLORKEY
'surface.BltFast X, Y, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, 1 Or 32 ' Or 16 'DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
surface.BltFast X, Y, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
'DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub



Sub DrawBackBufferSurface()
Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT)

End Sub
Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function
Sub DrawGrhtoHdc(Hwnd As Long, hDC As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)

If Grh <= 0 Then Exit Sub

SecundaryClipper.SetHWnd Hwnd
SurfaceDB.GetBMP(GrhData(Grh).FileNum).BltToDC hDC, SourceRect, destRect

End Sub
Sub PlayWaveAPI(File As String)
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub

Function EsConsejoReal(Nombre As String) As Boolean
'If UCase$(Nombre) <> "VELMONTH" And _
UCase$(Nombre) <> "BURROWS" And _
UCase$(Nombre) <> "BRIGHTSIDE" And _
UCase$(Nombre) <> "DREAMWEAVER" And _
UCase$(Nombre) <> "MISERY" Then Exit Function
'EsConsejoReal = True
End Function
Function EsConsejoCaoz(Nombre As String) As Boolean
'If UCase$(Nombre) <> "TIFFANY" And _
UCase$(Nombre) <> "ROMANA" And _
UCase$(Nombre) <> "PHOE" And _
UCase$(Nombre) <> "RA" And _
UCase$(Nombre) <> "FISH" Then Exit Function
'EsConsejoCaoz = True

End Function



Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
On Error Resume Next
If UserCiego Then Exit Sub
Dim Y        As Integer
Dim X        As Integer
Dim minY     As Integer
Dim maxY     As Integer
Dim minX     As Integer
Dim maxX     As Integer
Dim ScreenX  As Integer
Dim ScreenY  As Integer
Dim Moved    As Byte
Dim Grh      As Grh
Dim tempchar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer
Dim iPPy     As Integer
Dim rSourceRect      As RECT
Dim iGrhIndex        As Integer
Dim PixelOffsetXTemp As Integer
Dim PixelOffsetYTemp As Integer
Dim Tiempo As Double

minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)
Tiempo = GetTickCount
ScreenY = 8 + RenderMod.iImageSize



For Y = (minY + 8) + RenderMod.iImageSize To (maxY - 8) - RenderMod.iImageSize
    ScreenX = 8 + RenderMod.iImageSize
    For X = (minX + 8) + RenderMod.iImageSize To (maxX - 8) - RenderMod.iImageSize
        If X > 100 Or Y < 1 Then Exit For
        
    

        
        With MapData(X, Y).Graphic(1)
            If (.Started = 1) Then
                If (.SpeedCounter > 0) Then
                    .SpeedCounter = .SpeedCounter - 1
                    If (.SpeedCounter = 0) Then
                        .SpeedCounter = GrhData(.GrhIndex).Speed
                        .FrameCounter = .FrameCounter + 1
                        If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then .FrameCounter = 1
                    End If
                End If
            End If

            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With
        

        rSourceRect.Left = GrhData(iGrhIndex).sX
        rSourceRect.Top = GrhData(iGrhIndex).sY
        rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight
        Call BackBufferSurface.BltFast(((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), rSourceRect, DDBLTFAST_WAIT)

       If Not RenderMod.bNoCostas And MapData(X, Y).Graphic(2).GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(2), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y

ScreenY = 8 + RenderMod.iImageSize
For Y = (minY + 8) + RenderMod.iImageSize To (maxY - 1) - RenderMod.iImageSize
    ScreenX = 5 + RenderMod.iImageSize
    For X = (minX + 5) + RenderMod.iImageSize To (maxX - 5) - RenderMod.iImageSize
        If Not (X > 100 Or X < 1) Then
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
    
            If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).ObjGrh, iPPx, iPPy, 1, 1)
            If MapData(X, Y).CharIndex > 0 Then
                tempchar = CharList(MapData(X, Y).CharIndex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                Moved = 0

            If tempchar.MoveOffset.X <> 0 Then
                tempchar.Body.Walk(tempchar.Heading).Started = 1
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 1
                tempchar.escudo.ShieldWalk(tempchar.Heading).Started = 1
                PixelOffsetXTemp = PixelOffsetXTemp + tempchar.MoveOffset.X
                tempchar.MoveOffset.X = tempchar.MoveOffset.X - (8 * Sgn(tempchar.MoveOffset.X))
                Moved = 1
            End If

            If tempchar.MoveOffset.Y <> 0 Then
                tempchar.Body.Walk(tempchar.Heading).Started = 1
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 1
                tempchar.escudo.ShieldWalk(tempchar.Heading).Started = 1
                PixelOffsetYTemp = PixelOffsetYTemp + tempchar.MoveOffset.Y
                tempchar.MoveOffset.Y = tempchar.MoveOffset.Y - (8 * Sgn(tempchar.MoveOffset.Y))
                Moved = 1
            End If

            If Moved = 0 And tempchar.Moving = 1 Then
                tempchar.Moving = 0
                tempchar.Body.Walk(tempchar.Heading).FrameCounter = 1
                tempchar.Body.Walk(tempchar.Heading).Started = 0
                tempchar.arma.WeaponWalk(tempchar.Heading).FrameCounter = 1
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 0
                tempchar.escudo.ShieldWalk(tempchar.Heading).FrameCounter = 1
                tempchar.escudo.ShieldWalk(tempchar.Heading).Started = 0
                tempchar.haciendoataque = 0
            End If
            
            If tempchar.haciendoataque = 0 And tempchar.MoveOffset.X = 0 And tempchar.MoveOffset.Y = 0 Then
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 0
                tempchar.arma.WeaponWalk(tempchar.Heading).FrameCounter = 1
                End If
            If tempchar.haciendoataque = 1 Then
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 1
                tempchar.haciendoataque = 0
            End If
            
                iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
                iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
                
                If Len(tempchar.Nombre) = 0 Then
                'Call BackBufferSurface.BltFast(((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, SurfaceDB.GetBMP(3500), rSourceRect, DDBLTFAST_WAIT)
                    Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                
                    If tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Head.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.X, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                Else
                    If tempchar.Navegando = 1 Then
                        Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                    ElseIf Not CharList(MapData(X, Y).CharIndex).invisible And tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                       ' Call DDrawBlendGrhtoSurface(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                        If tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Head.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.X, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                        
                        If tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(SupBMiniMap, tempchar.Head.Head(tempchar.Heading), 32, 32, 1, 0)
                        
                        
                        If tempchar.casco.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.casco.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.X, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                        If tempchar.arma.WeaponWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.arma.WeaponWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                        If tempchar.escudo.ShieldWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.escudo.ShieldWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                        
                        
             
                    End If
                       ' Exit Sub
                    If Nombres Then
                        
                          If Not (tempchar.invisible Or tempchar.Navegando = 1) Then
        
                            Dim lCenter As Long
                            If InStr(tempchar.Nombre, "<") > 0 And InStr(tempchar.Nombre, ">") > 0 Then
                                Dim sClan As String
                                If BancoT > 0 And tempchar.Nombre = CharList(UserCharIndex).Nombre Then
                                    lCenter = (frmMain.TextWidth(Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1)) / 2) - 16
                                    sClan = Mid$(tempchar.Nombre, InStr(tempchar.Nombre, "<"))
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1), RGB(127, 0, 0))
                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(127, 0, 0))
                                Else
                                    lCenter = (frmMain.TextWidth(Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1)) / 2) - 16
                                    sClan = Mid$(tempchar.Nombre, InStr(tempchar.Nombre, "<"))
                                     If ModoGuerra = True And (InStr(CharList(UserCharIndex).Nombre, "<") > 0 And InStr(CharList(UserCharIndex).Nombre, ">") > 0) Then
                                    If sClan = Mid$(CharList(UserCharIndex).Nombre, InStr(CharList(UserCharIndex).Nombre, "<")) Then
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1), RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                                        lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(255, 128, 64))
                                    Else
                    
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1), RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                                        lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                                                    
                                    End If
                                    Else
                                    
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1), RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                                        lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                        Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                                    End If
                                End If
                            Else
                                lCenter = (frmMain.TextWidth(tempchar.Nombre) / 2) - 16
                                If BancoT > 0 And tempchar.Nombre = CharList(UserCharIndex).Nombre Then
                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempchar.Nombre, RGB(127, 0, 0))
                                Else
                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempchar.Nombre, RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                                End If
                            End If
                        
                        
                          Else
                          
                          
                          Dim UserClan As String
                          UserClan = Mid$(CharList(UserCharIndex).Nombre, InStr(CharList(UserCharIndex).Nombre, "<"))
 '
                            'si tengo clan
                            If UserClan <> "" Then
                                sClan = ""
                                sClan = Mid$(tempchar.Nombre, InStr(tempchar.Nombre, "<"))
                                If sClan = UserClan Then
                                    lCenter = (frmMain.TextWidth(Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1)) / 2) - 16
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1), vbYellow)
                                    lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, vbYellow)
                                    Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                                    If tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.Head.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.X, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                                    If tempchar.casco.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.casco.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.X, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                                    If tempchar.arma.WeaponWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.arma.WeaponWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                                    If tempchar.escudo.ShieldWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.escudo.ShieldWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                                End If
                                'SI ES MI CLAN
                            Else
                                'SISOY YO MISMO
                                If CharList(UserCharIndex).Nombre = tempchar.Nombre Then
                                    lCenter = (frmMain.TextWidth(tempchar.Nombre) / 2) - 16
                                    Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempchar.Nombre, vbYellow)
                                    Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                                    If tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.Head.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.X, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                                    If tempchar.casco.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.casco.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.X, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                                    If tempchar.arma.WeaponWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.arma.WeaponWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                                    If tempchar.escudo.ShieldWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, tempchar.escudo.ShieldWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                                End If
                                '/
                                
                            End If
                        
                        
                        If UCase$(UserClan) = "<FURIUSAO STAFF>" Or UCase$(UserClan) = "<CONSEJERO>" And Not CharList(UserCharIndex).Nombre = tempchar.Nombre Then
                                      lCenter = (frmMain.TextWidth(tempchar.Nombre) / 2) - 16
                                      Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempchar.Nombre, vbYellow)
                        End If
                        
                        
                        End If
                       
                    End If
                End If
    
                If Dialogos.CantidadDialogos > 0 Then Call Dialogos.Update_Dialog_Pos((iPPx + tempchar.Body.HeadOffset.X), (iPPy + tempchar.Body.HeadOffset.Y), MapData(X, Y).CharIndex)
                
                CharList(MapData(X, Y).CharIndex) = tempchar
    
                If CharList(MapData(X, Y).CharIndex).FX <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, FxData(tempchar.FX).FX, iPPx + FxData(tempchar.FX).OffsetX, iPPy + FxData(tempchar.FX).OffsetY, 1, 1, MapData(X, Y).CharIndex)
                
            End If

        End If
        If MapData(X, Y).Graphic(3).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(3), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
        ' If MapData(X, Y).Graphic(3).GrhIndex > 0 Then MsgBox MapData(X, Y).Graphic(3).GrhIndex
        'Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(3), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y
'Exit Sub
If Not bTecho Then
    ScreenY = 5 + RenderMod.iImageSize
    For Y = (minY + 5) + RenderMod.iImageSize To (maxY - 1) - RenderMod.iImageSize
        ScreenX = 5 + RenderMod.iImageSize
        For X = (minX + 5) + RenderMod.iImageSize To (maxX - 0) - RenderMod.iImageSize
            
            If X < 101 And X > 0 And Y < 101 And Y > 0 Then
                If frmOpciones.TransTecho.Value Then
                    If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then Call DDrawTransGrhtoSurfaceAlpha(BackBufferSurface, MapData(X, Y).Graphic(4), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
                Else
                    If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(4), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
                End If
            End If
            
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
End If
Exit Sub

If bLluvia(UserMap) = 1 Then
    If bRain Or bRainST Then
        If llTick < DirectX.TickCount - 50 Then
            iFrameIndex = iFrameIndex + 1
            If iFrameIndex > 7 Then iFrameIndex = 0
            llTick = DirectX.TickCount
        End If

        For Y = 0 To 4
            For X = 0 To 4
                'Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(X), SurfaceDB.GetBMP(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
            Next X
        Next Y
    End If
End If
'Lieto nieve


'If bLluvia(UserMap) = 1 Then
    If bSnow Or bSnowST Then
        If llTick < DirectX.TickCount - 50 Then
            iFrameIndex = iFrameIndex + 1
            If iFrameIndex > 7 Then iFrameIndex = 0
            llTick = DirectX.TickCount
        End If

        For Y = 0 To 4
            For X = 0 To 4
                'Call BackBufferSurface.BltFast(LTNieva(Y), LTNieva(X), SurfaceDB.GetBMP(5557), RNieva(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
            Next X
        Next Y
    End If
'End If
'Leito nieva
'


End Sub


Public Function RenderSounds()

    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> plLluviain Then
                    Call frmMain.StopSound
                    Call frmMain.Play("lluviain.wav", True)
                    frmMain.IsPlaying = plLluviain
                End If
                
                
            Else
                If frmMain.IsPlaying <> plLluviaout Then
                    Call frmMain.StopSound
                    Call frmMain.Play("lluviaout.wav", True)
                    frmMain.IsPlaying = plLluviaout
                End If
                
                
            End If
        End If
    End If

End Function


Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            CharList(UserCharIndex).POS.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).POS.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).POS.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And CharList(UserCharIndex).POS.Y <= Y
        
End If

End Function



Function PixelPos(X As Integer) As Integer




PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function


Sub LoadGraphics()
        Dim loopc As Integer
        Dim SurfaceDesc As DDSURFACEDESC2
        Dim ddck As DDCOLORKEY
        Dim ddsd As DDSURFACEDESC2
        Dim iLoopUpdate As Integer

        SurfaceDB.MaxEntries = 400
        SurfaceDB.lpDirectDraw7 = DirectDraw
        SurfaceDB.Path = DirGraficos
        Call SurfaceDB.Init(mododinamico)

        
        Call GetBitmapDimensions(DirGraficos & "5556.bmp", ddsd.lWidth, ddsd.lHeight)
              
        RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
        RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
        RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
        RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    
        RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
        RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
        RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
        RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
        AddtoRichTextBox frmCargando.Status, "Hecho.", 255, 150, 50, 1, , False
         'Letio NIEVE (Bmp de la nieve)
        Call GetBitmapDimensions(DirGraficos & "5557.bmp", ddsd.lWidth, ddsd.lHeight)
        RNieva(0).Top = 0:      RNieva(1).Top = 0:      RNieva(2).Top = 0:      RNieva(3).Top = 0
        RNieva(0).Left = 0:     RNieva(1).Left = 128:   RNieva(2).Left = 256:   RNieva(3).Left = 384
        RNieva(0).Right = 128:  RNieva(1).Right = 256:  RNieva(2).Right = 384:  RNieva(3).Right = 512
        RNieva(0).Bottom = 128: RNieva(1).Bottom = 128: RNieva(2).Bottom = 128: RNieva(3).Bottom = 128
        RNieva(4).Top = 128:    RNieva(5).Top = 128:    RNieva(6).Top = 128:    RNieva(7).Top = 128
        RNieva(4).Left = 0:     RNieva(5).Left = 128:   RNieva(6).Left = 256:   RNieva(7).Left = 384
        RNieva(4).Right = 128:  RNieva(5).Right = 256:  RNieva(6).Right = 384:  RNieva(7).Right = 512
        RNieva(4).Bottom = 256: RNieva(5).Bottom = 256: RNieva(6).Bottom = 256: RNieva(7).Bottom = 256
        AddtoRichTextBox frmCargando.Status, "Hecho.", , , , 1, , False
End Sub



Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean





Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"


UserPos.X = MinXBorder
UserPos.Y = MinYBorder



DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

If Musica = 0 Or FX = 0 Then
    DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
End If



With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.Hwnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0 + 32 * RenderMod.iImageSize
    .Top = 0 + 32 * RenderMod.iImageSize
    .Right = (TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))) - 32 * RenderMod.iImageSize
    .Bottom = (TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))) - 32 * RenderMod.iImageSize
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If RenderMod.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)


'BufferPJxD
ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck

Call InitBlend(BackBufferSurface)

Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs


LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

'Leito nieve
LTNieva(0) = 224
LTNieva(1) = 352
LTNieva(2) = 480
LTNieva(3) = 608
LTNieva(4) = 736
'Leito nieve

AddtoRichTextBox frmCargando.Status, "Cargando Gráficos....", 255, 150, 50, , , True
Call LoadGraphics

InitTileEngine = True
End Function





Sub ShowNextFrame()












    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer

    If EngineRun Then
        
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = (OffsetCounterX - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.X)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            

            
            
            ElseIf AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If

            
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
          '  If ModoTrabajo Then Call Dialogos.DrawText(260, 260, "MODO TRABAJO", vbRed)
            
            Call DibujarCartel
            
            Call DrawBackBufferSurface
            
            
            FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, Index As Integer)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1)
Call CrearGrh(534, 2)
End Sub

