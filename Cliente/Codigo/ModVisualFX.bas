Attribute VB_Name = "ModVisualFX"


Option Explicit

Public icMode As Integer
Public UseAlphaBlending As Boolean

'Public Declare Function AlphaBlend Lib "AoFX.dll" (ByVal iMode As Integer, ByVal bColorKey As Integer, ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer
Public Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

Public Sub EfectoNoche(ByRef surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim modo As Long
Dim rRect As RECT

surface.GetSurfaceDesc ddsdDest

With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    modo = 1
Else
    modo = 2
End If

Dim DstLock As Boolean
DstLock = False

On Local Error GoTo HayErrorAlpha

surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

surface.GetLockedArray dArray()
Call BltEfectoNoche(ByVal VarPtr(dArray(0, 0)), _
    ddsdDest.lWidth, ddsdDest.lHeight, ddsdDest.lPitch, _
    modo)
    
HayErrorAlpha:

If DstLock = True Then
    surface.Unlock rRect
    DstLock = False
End If

End Sub

Public Sub EfectoTarde(ByRef surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim modo As Long
Dim rRect As RECT

surface.GetSurfaceDesc ddsdDest

With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    modo = 1
Else
    modo = 2
End If

Dim DstLock As Boolean
DstLock = False

On Local Error GoTo HayErrorAlpha

surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

surface.GetLockedArray dArray()

    
Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 60, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 10, 10, 0)
    
HayErrorAlpha:

If DstLock = True Then
    surface.Unlock rRect
    DstLock = False
End If

End Sub

Public Sub EfectoAmanecer(ByRef surface As DirectDrawSurface7)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim modo As Long
Dim rRect As RECT

surface.GetSurfaceDesc ddsdDest

With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With

If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
    modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
    modo = 1
Else
    modo = 2
End If

Dim DstLock As Boolean
DstLock = False

On Local Error GoTo HayErrorAlpha

surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

surface.GetLockedArray dArray()

    
Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 80, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 0, 50, 50)
HayErrorAlpha:

If DstLock = True Then
    surface.Unlock rRect
    DstLock = False
End If
'Code by Ladder 11/12/07

End Sub



Sub InitBlend(surface As DirectDrawSurface7)
If UseAlphaBlending Then
    Dim ddsdtemp As DDSURFACEDESC2
          Call surface.GetSurfaceDesc(ddsdtemp)
          
          Select Case ddsdtemp.ddpfPixelFormat.lGBitMask
            Case &H3E0
                icMode = 555
            Case &H7E0
                icMode = 565
            Case Else
                MsgBox "No se pudo detectar el modo del BackBuffer ¿Esta en 16 bits de colores?"
                UseAlphaBlending = False
          End Select
End If
End Sub

   Sub DDrawTransGrhtoSurfaceAlpha(surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'If Screen.Width / Screen.TwipsPerPixelY > 800 Then X = X + (1024 - 800) * 2 + 10
'[END]'
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
'[CODE]:MatuX
'
'  CurrentGrh.GrhIndex = iGrhIndex
'
'[END]

'Dim CurrentGrh As Grh
Dim iGrhIndex As Integer
'Dim destRect As RECT
Dim SourceRect As RECT
'Dim SurfaceDesc As DDSURFACEDESC2
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
               End If
            End If
        End If
    End If
End If

'If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

'Center Grh over X,Y pos
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

'surface.BltFast X, Y, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim modo As Long

Set Src = SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum)

Src.GetSurfaceDesc ddsdSrc
surface.GetSurfaceDesc ddsdDest

With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
    
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With

' 0 -> 16 bits 555
' 1 -> 16 bits 565
' 2 -> 16 bits raro (Sin implementar)
' 3 -> 24 bits
' 4 -> 32 bits
'
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H3E0 Then
    modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    modo = 1
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 And ddsdSrc.ddpfPixelFormat.lGBitMask = &H7E0 Then
    modo = 3
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = 65280 And ddsdSrc.ddpfPixelFormat.lGBitMask = 65280 Then
    modo = 4
Else
    'Modo = 2 '16 bits raro ?
  '  surface.BltFast X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Exit Sub
End If

Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False

On Local Error GoTo HayErrorAlpha

Src.Lock SourceRect, ddsdSrc, DDLOCK_WAIT, 0
SrcLock = True
surface.Lock rDest, ddsdDest, DDLOCK_WAIT, 0
DstLock = True

surface.GetLockedArray dArray()
Src.GetLockedArray sArray()

Call BltAlphaFast(ByVal VarPtr(dArray(X + X, Y)), ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, modo)

surface.Unlock rDest
DstLock = False
Src.Unlock SourceRect
SrcLock = False


Exit Sub

HayErrorAlpha:
If SrcLock Then Src.Unlock SourceRect
If DstLock Then surface.Unlock rDest
surface.BltFast X, Y, Src, SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub


