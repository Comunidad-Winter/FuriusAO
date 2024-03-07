Attribute VB_Name = "DibujarInventario"


Option Explicit







Public Const XCantItems = 5

Public OffsetDelInv As Integer
Public ItemElegido As Integer
Public mx As Integer
Public my As Integer

Private AuxSurface   As DirectDrawSurface7
Private BoxSurface   As DirectDrawSurface7
Private SelSurface   As DirectDrawSurface7
Private bStaticInit  As Boolean
Private r1           As RECT, r2 As RECT, auxr As RECT
Private rBox         As RECT
Private rBoxFrame(2) As RECT
Private iFrameMod    As Integer
Sub ActualizarOtherInventory(Slot As Integer)

If OtherInventory(Slot).OBJIndex = 0 Then
    frmComerciar.List1(0).List(Slot - 1) = "Nada"
Else
    frmComerciar.List1(0).List(Slot - 1) = OtherInventory(Slot).Name
End If

If frmComerciar.List1(0).ListIndex = Slot - 1 And lista = 0 Then Call ActualizarInformacionComercio(0)

End Sub
Sub ActualizarInventario(Slot As Integer)
Dim OBJIndex As Long
Dim NameSize As Byte

If UserInventory(Slot).Amount = 0 Then
    frmMain.imgObjeto(Slot).ToolTipText = "Nada"
    frmMain.lblObjCant(Slot).ToolTipText = "Nada"
    frmMain.lblObjCant(Slot).Caption = ""
    FormInv.imgObjeto(Slot).ToolTipText = "Nada"
    FormInv.lblObjCant(Slot).ToolTipText = "Nada"
    FormInv.lblObjCant(Slot).Caption = ""
    If ItemElegido = Slot Then frmMain.Shape1.Visible = False
    If ItemElegido = Slot Then FormInv.Shape1.Visible = False
Else
    frmMain.imgObjeto(Slot).ToolTipText = UserInventory(Slot).Name
    frmMain.lblObjCant(Slot).ToolTipText = UserInventory(Slot).Name
    frmMain.lblObjCant(Slot).Caption = CStr(UserInventory(Slot).Amount)
    FormInv.imgObjeto(Slot).ToolTipText = UserInventory(Slot).Name
    FormInv.lblObjCant(Slot).ToolTipText = UserInventory(Slot).Name
    FormInv.lblObjCant(Slot).Caption = CStr(UserInventory(Slot).Amount)

    If ItemElegido = Slot Then frmMain.Shape1.Visible = True
    If ItemElegido = Slot Then FormInv.Shape1.Visible = True
End If



If UserInventory(Slot).GrhIndex > 0 Then

    
    frmMain.imgObjeto(Slot).Picture = LoadPicture(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".bmp")
    FormInv.imgObjeto(Slot).Picture = LoadPicture(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".bmp")


Else
    frmMain.imgObjeto(Slot).Picture = LoadPicture()
    FormInv.imgObjeto(Slot).Picture = LoadPicture()

End If

If UserInventory(Slot).Equipped > 0 Then
    frmMain.Label2(Slot).Visible = True
    FormInv.Label2(Slot).Visible = True
Else
    frmMain.Label2(Slot).Visible = False
    FormInv.Label2(Slot).Visible = False
End If

If frmComerciar.Visible Then
    If UserInventory(Slot).Amount = 0 Then
        frmComerciar.List1(1).List(Slot - 1) = "Nada"
     Else
        frmComerciar.List1(1).List(Slot - 1) = UserInventory(Slot).Name
    End If
    If frmComerciar.List1(1).ListIndex = Slot - 1 And lista = 1 Then Call ActualizarInformacionComercio(1)
End If

End Sub
Private Sub InitMem()
    Dim ddck        As DDCOLORKEY
    Dim SurfaceDesc As DDSURFACEDESC2
    
    
    r1.Right = 32: r1.Bottom = 32
    r2.Right = 32: r2.Bottom = 32
    
    With SurfaceDesc
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lHeight = r1.Bottom
        .lWidth = r1.Right
    End With

    
    Set AuxSurface = DirectDraw.CreateSurface(SurfaceDesc)
    Set BoxSurface = DirectDraw.CreateSurface(SurfaceDesc)
    Set SelSurface = DirectDraw.CreateSurface(SurfaceDesc)

    
    AuxSurface.SetColorKey DDCKEY_SRCBLT, ddck
    BoxSurface.SetColorKey DDCKEY_SRCBLT, ddck
    SelSurface.SetColorKey DDCKEY_SRCBLT, ddck

    auxr.Right = 32: auxr.Bottom = 32

    AuxSurface.SetFontTransparency True
    AuxSurface.SetFont frmMain.Font
    SelSurface.SetFontTransparency True
    SelSurface.SetFont frmMain.Font

    
    With rBoxFrame(0): .Left = 0:  .Top = 0: .Right = 32: .Bottom = 32: End With
    With rBoxFrame(1): .Left = 32: .Top = 0: .Right = 64: .Bottom = 32: End With
    With rBoxFrame(2): .Left = 64: .Top = 0: .Right = 96: .Bottom = 32: End With
    iFrameMod = 1

    bStaticInit = True
End Sub
