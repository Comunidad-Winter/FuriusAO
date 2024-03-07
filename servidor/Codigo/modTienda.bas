Attribute VB_Name = "modTienda"

Public DineroTotalVentas As Double
Public NumeroVentas As Long

Option Explicit
Sub TiendaVentaItem(userindex As Integer, ByVal i As Integer, Cantidad As Integer, NpcIndex As Integer)
On Error GoTo errhandler
Dim Vendedor As Integer

If Cantidad < 1 Or Npclist(NpcIndex).NPCtype <> NPCTYPE_TIENDA Then Exit Sub

Vendedor = Npclist(NpcIndex).flags.TiendaUser

If UserList(userindex).Stats.GLD >= (UserList(Vendedor).Tienda.Object(i).Precio * Cantidad) Then
    If UserList(Vendedor).Tienda.Object(i).Amount Then
         If Cantidad > UserList(Vendedor).Tienda.Object(i).Amount Then Cantidad = UserList(Vendedor).Tienda.Object(i).Amount
         Call TiendaCompraItem(userindex, CInt(i), UserList(userindex).flags.TargetNpc, Cantidad)
         Call SendUserORO(userindex)
    Else
        Call SendData(ToIndex, userindex, 0, "OTIV" & i)
    End If
Else
    Call SendData(ToIndex, userindex, 0, "2Q")
    Exit Sub
End If

errhandler:

End Sub
Sub TiendaCompraItem(userindex As Integer, Slot As Byte, NpcIndex As Integer, Cantidad As Integer)
Dim Vendedor As Integer
Dim ObjI As Integer
Dim Encontre As Boolean
Dim MiObj As Obj

Vendedor = Npclist(NpcIndex).flags.TiendaUser

If (UserList(Vendedor).Tienda.Object(Slot).Amount <= 0) Then Exit Sub

ObjI = UserList(Vendedor).Tienda.Object(Slot).OBJIndex

MiObj.OBJIndex = ObjI
MiObj.Amount = Cantidad

If Not MeterItemEnInventario(userindex, MiObj) Then
    Call SendData(ToIndex, userindex, 0, "5P")
    Exit Sub
End If

UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad

Call VendedorVenta(Vendedor, CByte(Slot), Cantidad, userindex)

End Sub
Sub VendedorVenta(Vendedor As Integer, Slot As Byte, Cantidad As Integer, Comprador As Integer)

Call SendData(ToIndex, Vendedor, 0, "/R" & UserList(Comprador).Name & "," & ObjData(UserList(Vendedor).Tienda.Object(Slot).OBJIndex).Name & "," & Cantidad & "," & UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad)
UserList(Vendedor).Stats.Banco = UserList(Vendedor).Stats.Banco + UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad
UserList(Vendedor).Tienda.Gold = UserList(Vendedor).Tienda.Gold + UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad
DineroTotalVentas = DineroTotalVentas + UserList(Vendedor).Tienda.Object(Slot).Precio * Cantidad
NumeroVentas = NumeroVentas + 1

UserList(Vendedor).Tienda.Object(Slot).Amount = UserList(Vendedor).Tienda.Object(Slot).Amount - Cantidad

If UserList(Vendedor).Tienda.Object(Slot).Amount <= 0 Then
    UserList(Vendedor).Tienda.Object(Slot).Amount = 0
    UserList(Vendedor).Tienda.Object(Slot).OBJIndex = 0
    UserList(Vendedor).Tienda.Object(Slot).Precio = 0
    UserList(Vendedor).Tienda.NroItems = UserList(Vendedor).Tienda.NroItems - 1
    If UserList(Vendedor).Tienda.NroItems <= 0 Then
        Npclist(UserList(Vendedor).Tienda.NpcTienda).flags.TiendaUser = 0
        UserList(Vendedor).Tienda.NpcTienda = 0
        Call SendData(ToIndex, Vendedor, 0, "/S")
        Call SendData(ToIndex, Comprador, 0, "FINCOMOK")
        Exit Sub
    End If
End If

Call UpdateTiendaC(False, Comprador, UserList(Vendedor).Tienda.NpcTienda, Slot)

Exit Sub
errhandler:

End Sub
Sub IniciarComercioTienda(userindex As Integer, NpcIndex As Integer)

Call UpdateTiendaC(True, userindex, NpcIndex, 0)
Call SendData(ToIndex, userindex, 0, "INITCOM")
UserList(userindex).flags.Comerciando = True

End Sub
Public Sub IniciarAlquiler(userindex As Integer)

If Not (ClaseTrabajadora(UserList(userindex).Clase) And Not EsNewbie(userindex)) And Not (UserList(userindex).Stats.ELV >= 25 And UserList(userindex).Stats.UserSkills(Comerciar) >= 65) Then
    Call SendData(ToIndex, userindex, 0, "/V" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(userindex).Tienda.NpcTienda > 0 And UserList(userindex).Tienda.NpcTienda <> UserList(userindex).flags.TargetNpc Then
    Call SendData(ToIndex, userindex, 0, "/W" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

Call UpdateTiendaV(True, userindex, 0)
Call SendData(ToIndex, userindex, 0, "INITIENDA")
UserList(userindex).flags.Comerciando = True

End Sub
Sub UpdateTiendaV(ByVal UpdateAll As Boolean, userindex As Integer, Slot As Byte, Optional ByVal TodaInfo As Boolean)
Dim i As Byte

If UpdateAll Then
    For i = 1 To MAX_TIENDA_SLOTS
        Call SendTiendaItemV(userindex, i, UpdateAll)
    Next
Else
    Call SendTiendaItemV(userindex, Slot, TodaInfo)
End If

End Sub
Sub SendTiendaItemV(userindex As Integer, Slot As Byte, TodaInfo As Boolean)
Dim MiObj As TiendaObj

MiObj = UserList(userindex).Tienda.Object(Slot)

If MiObj.OBJIndex Then
    If TodaInfo Then
        Call SendData(ToIndex, userindex, 0, "OTII" & Slot _
        & "," & ObjData(MiObj.OBJIndex).Name _
        & "," & MiObj.Amount _
        & "," & MiObj.Precio _
        & "," & ObjData(MiObj.OBJIndex).GrhIndex _
        & "," & MiObj.OBJIndex _
        & "," & ObjData(MiObj.OBJIndex).ObjType _
        & "," & ObjData(MiObj.OBJIndex).MaxHIT _
        & "," & ObjData(MiObj.OBJIndex).MinHIT _
        & "," & ObjData(MiObj.OBJIndex).MaxDef _
        & "," & ObjData(MiObj.OBJIndex).MinDef _
        & "," & ObjData(MiObj.OBJIndex).TipoPocion _
        & "," & ObjData(MiObj.OBJIndex).MaxModificador _
        & "," & ObjData(MiObj.OBJIndex).MinModificador)
    Else
        Call SendData(ToIndex, userindex, 0, "OTIC " & Slot & "," & MiObj.Amount)
    End If
Else
    Call SendData(ToIndex, userindex, 0, "OTIV" & Slot)
End If

End Sub
Sub UpdateTiendaC(ByVal UpdateAll As Boolean, userindex As Integer, NpcIndex As Integer, Slot As Byte)
Dim i As Byte

If UpdateAll Then
    For i = 1 To MAX_TIENDA_SLOTS
        Call SendTiendaItemC(userindex, NpcIndex, i, UpdateAll)
    Next
Else
    Call SendTiendaItemC(userindex, NpcIndex, Slot, UpdateAll)
End If

End Sub
Sub SendTiendaItemC(userindex As Integer, NpcIndex As Integer, Slot As Byte, ByVal TodaInfo As Boolean)
Dim MiObj As TiendaObj

MiObj = UserList(Npclist(NpcIndex).flags.TiendaUser).Tienda.Object(Slot)

If MiObj.OBJIndex Then
    If TodaInfo Then
        Call SendData(ToIndex, userindex, 0, "OTII" & Slot _
        & "," & ObjData(MiObj.OBJIndex).Name _
        & "," & MiObj.Amount _
        & "," & MiObj.Precio _
        & "," & ObjData(MiObj.OBJIndex).GrhIndex _
        & "," & MiObj.OBJIndex _
        & "," & ObjData(MiObj.OBJIndex).ObjType _
        & "," & ObjData(MiObj.OBJIndex).MaxHIT _
        & "," & ObjData(MiObj.OBJIndex).MinHIT _
        & "," & ObjData(MiObj.OBJIndex).MaxDef _
        & "," & ObjData(MiObj.OBJIndex).MinDef _
        & "," & ObjData(MiObj.OBJIndex).TipoPocion _
        & "," & ObjData(MiObj.OBJIndex).MaxModificador _
        & "," & ObjData(MiObj.OBJIndex).MinModificador _
        & "," & PuedeUsarObjeto(userindex, MiObj.OBJIndex))
    Else
        Call SendData(ToIndex, userindex, 0, "OTIC" & Slot & "," & MiObj.Amount)
    End If
Else
    Call SendData(ToIndex, userindex, 0, "OTIV" & Slot)
End If

End Sub
Sub UserSacaVenta(userindex As Integer, Slot As Byte, Cantidad As Integer)
On Error GoTo errhandler

If UserList(userindex).Tienda.Object(Slot).Amount Then
    If Cantidad > UserList(userindex).Tienda.Object(Slot).Amount Then Cantidad = UserList(userindex).Tienda.Object(Slot).Amount
    Call UserSacaObjVenta(userindex, CInt(Slot), Cantidad)
End If

Exit Sub
errhandler:

End Sub
Sub UserPoneVenta(userindex As Integer, Slot As Byte, Cantidad As Integer, Precio As Long)
On Error GoTo errhandler

If ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex).Newbie Then
    Call SendData(ToIndex, userindex, 0, "/H")
    Exit Sub
End If

If ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex).NoSeCae Then
    Call SendData(ToIndex, userindex, 0, "||No puedes poner este objeto a la venta." & FONTTYPE_INFO)
    Exit Sub
End If

If ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex).Caos > 0 Or ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex).Real Then
    Call SendData(ToIndex, userindex, 0, "/I")
    Exit Sub
End If

If Precio = 0 Then
    Call SendData(ToIndex, userindex, 0, "/M")
    Exit Sub
ElseIf Precio >= 100000 Then
    Call SendData(ToIndex, userindex, 0, "||El precio debe ser menor a 100000 monedas de oro" & FONTTYPE_BLANCO)
    Exit Sub
End If


If UserList(userindex).Tienda.NpcTienda = 0 Then
    UserList(userindex).Tienda.NpcTienda = UserList(userindex).flags.TargetNpc
    Npclist(UserList(userindex).flags.TargetNpc).flags.TiendaUser = userindex
End If

If UserList(userindex).Invent.Object(Slot).Amount > 0 And UserList(userindex).Invent.Object(Slot).Equipped = 0 Then
    If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Slot).Amount Then Cantidad = UserList(userindex).Invent.Object(Slot).Amount
    Call UserDaObjVenta(userindex, CInt(Slot), Cantidad, Precio)
End If

Exit Sub
errhandler:

End Sub
Sub UserSacaObjVenta(userindex As Integer, ByVal Itemslot As Byte, Cantidad As Integer)
Dim MiObj As Obj

If Cantidad < 1 Then Exit Sub

MiObj.OBJIndex = UserList(userindex).Tienda.Object(Itemslot).OBJIndex
MiObj.Amount = Cantidad

If Not MeterItemEnInventario(userindex, MiObj) Then
    Call SendData(ToIndex, userindex, 0, "/J")
    Exit Sub
End If

UserList(userindex).Tienda.Object(Itemslot).Amount = UserList(userindex).Tienda.Object(Itemslot).Amount - Cantidad

If UserList(userindex).Tienda.Object(Itemslot).Amount <= 0 Then
    UserList(userindex).Tienda.Object(Itemslot).Amount = 0
    UserList(userindex).Tienda.Object(Itemslot).OBJIndex = 0
    UserList(userindex).Tienda.Object(Itemslot).Precio = 0
    UserList(userindex).Tienda.NroItems = UserList(userindex).Tienda.NroItems - 1
    If UserList(userindex).Tienda.NroItems <= 0 Then
        Npclist(UserList(userindex).Tienda.NpcTienda).flags.TiendaUser = 0
        UserList(userindex).Tienda.NpcTienda = 0
    End If
End If

Call UpdateTiendaV(False, userindex, Itemslot)

End Sub
Sub UserDaObjVenta(userindex As Integer, ByVal Itemslot As Byte, Cantidad As Integer, ByVal Precio As Long)
Dim Slot As Byte
Dim ObjI As Integer
Dim SlotHayado As Boolean

If Cantidad < 1 Then Exit Sub

ObjI = UserList(userindex).Invent.Object(Itemslot).OBJIndex
    
'For Slot = 1 To MAX_TIENDA_SLOTS
'    If UserList(userindex).Tienda.Object(Slot).OBJIndex = ObjI Then
'        SlotHayado = True
'        Exit For
'    End If
'Next

If Not SlotHayado Then
    For Slot = 1 To MAX_TIENDA_SLOTS
        If UserList(userindex).Tienda.Object(Slot).OBJIndex = 0 Then
            If UserList(userindex).Tienda.NroItems + UserList(userindex).BancoInvent.NroItems + 1 > MAX_BANCOINVENTORY_SLOTS Then
                Call SendData(ToIndex, userindex, 0, "/K")
                Exit Sub
            End If
            UserList(userindex).Tienda.NroItems = UserList(userindex).Tienda.NroItems + 1
            SlotHayado = True
            Exit For
        End If
    Next
End If

If Not SlotHayado Then
    Call SendData(ToIndex, userindex, 0, "/G")
    Exit Sub
End If

If UserList(userindex).Tienda.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    UserList(userindex).Tienda.Object(Slot).OBJIndex = ObjI
    UserList(userindex).Tienda.Object(Slot).Amount = UserList(userindex).Tienda.Object(Slot).Amount + Cantidad
    UserList(userindex).Tienda.Object(Slot).Precio = Precio
    Call QuitarUserInvItem(userindex, CByte(Itemslot), Cantidad)
Else
    Call SendData(ToIndex, userindex, 0, "/G")
End If

Call UpdateUserInv(False, userindex, CByte(Itemslot))
Call UpdateTiendaV(False, userindex, Slot, True)

End Sub
Sub DevolverItemsVenta(userindex As Integer)
Dim i As Byte


For i = 1 To MAX_TIENDA_SLOTS
    If UserList(userindex).Tienda.Object(i).OBJIndex Then Call TiendaABoveda(userindex, i)
Next

End Sub
Sub TiendaABoveda(userindex As Integer, Itemslot As Byte)
Dim Slot As Byte
Dim ObjI As Integer
Dim SlotHayado As Boolean

ObjI = UserList(userindex).Tienda.Object(Itemslot).OBJIndex
    
For Slot = 1 To MAX_BANCOINVENTORY_SLOTS
    If UserList(userindex).BancoInvent.Object(Slot).OBJIndex = ObjI Then
        SlotHayado = True
        Exit For
    End If
Next

If Not SlotHayado Then
    For Slot = 1 To MAX_BANCOINVENTORY_SLOTS
        If UserList(userindex).BancoInvent.Object(Slot).OBJIndex = 0 Then
            UserList(userindex).BancoInvent.NroItems = UserList(userindex).BancoInvent.NroItems + 1
            SlotHayado = True
            Exit For
        End If
    Next
End If

If Not SlotHayado Then Exit Sub

If UserList(userindex).BancoInvent.Object(Slot).Amount + UserList(userindex).Tienda.Object(Itemslot).Amount <= MAX_INVENTORY_OBJS Then
    UserList(userindex).BancoInvent.Object(Slot).OBJIndex = ObjI
    UserList(userindex).BancoInvent.Object(Slot).Amount = UserList(userindex).BancoInvent.Object(Slot).Amount + UserList(userindex).Tienda.Object(Itemslot).Amount
    UserList(userindex).Tienda.Object(Itemslot).Amount = 0
    UserList(userindex).Tienda.Object(Itemslot).OBJIndex = 0
    UserList(userindex).Tienda.Object(Itemslot).Precio = 0
    UserList(userindex).Tienda.NroItems = UserList(userindex).Tienda.NroItems - 1
End If

End Sub
