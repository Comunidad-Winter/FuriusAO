Attribute VB_Name = "Comercio"

Option Explicit
Sub UserCompraObj(userindex As Integer, ByVal OBJIndex As Integer, NpcIndex As Integer, Cantidad As Integer)
Dim Infla As Integer
Dim Desc As Single
Dim unidad As Long, monto As Long
Dim Slot As Byte
Dim ObjI As Integer
Dim Encontre As Boolean

ObjI = Npclist(UserList(userindex).flags.TargetNpc).Invent.Object(OBJIndex).OBJIndex

Slot = 1
Do Until UserList(userindex).Invent.Object(Slot).OBJIndex = ObjI And _
   UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then Exit Do
Loop

If Slot > MAX_INVENTORY_SLOTS Then
    Slot = 1
    Do Until UserList(userindex).Invent.Object(Slot).OBJIndex = 0
        Slot = Slot + 1

        If Slot > MAX_INVENTORY_SLOTS Then
            Call SendData(ToIndex, userindex, 0, "5P")
            Exit Sub
        End If
    Loop
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If

If UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    UserList(userindex).Invent.Object(Slot).OBJIndex = ObjI
    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + Cantidad
    Infla = (Npclist(NpcIndex).Inflacion * ObjData(ObjI).Valor) \ 100

    Desc = Descuento(userindex)
    
    unidad = Int(((ObjData(Npclist(NpcIndex).Invent.Object(OBJIndex).OBJIndex).Valor + Infla) / Desc))
    If unidad = 0 Then unidad = 1
    monto = unidad * Cantidad
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - monto
    
    Call SubirSkill(userindex, Comerciar)
    
    If ObjData(ObjI).ObjType = OBJTYPE_LLAVES Then Call LogVentaCasa(UserList(userindex).Name & " compro " & ObjData(ObjI).Name)
    Call QuitarNpcInvItem(UserList(userindex).flags.TargetNpc, CByte(OBJIndex), Cantidad, userindex)
    
    Call UpdateUserInv(False, userindex, Slot)
Else
    Call SendData(ToIndex, userindex, 0, "5P")
End If

End Sub
Sub UpdateNPCInv(UpdateAll As Boolean, userindex As Integer, NpcIndex As Integer, Slot As Byte)
Dim i As Byte
Dim MiObj As UserOBJ

If UpdateAll Then
    For i = 1 To MAX_NPCINVENTORY_SLOTS
        Call SendNPCItem(userindex, NpcIndex, i, UpdateAll)
    Next
Else
    Call SendNPCItem(userindex, NpcIndex, i, UpdateAll)
End If

End Sub
Sub SendNPCItem(userindex As Integer, NpcIndex As Integer, Slot As Byte, ByVal AllInfo As Boolean)
Dim MiObj As UserOBJ
Dim Infla As Long
Dim Desc As Single
Dim val As Long

MiObj = Npclist(NpcIndex).Invent.Object(Slot)

Desc = Descuento(userindex)

If Desc >= 0 And Desc <= 1 Then Desc = 1




If MiObj.OBJIndex Then
    If AllInfo Then
        Infla = (Npclist(NpcIndex).Inflacion * ObjData(MiObj.OBJIndex).Valor) / 100
        val = Maximo(1, Int((ObjData(MiObj.OBJIndex).Valor + Infla) / Desc))
        Call SendData(ToIndex, userindex, 0, "OTII" & Slot _
        & "," & ObjData(MiObj.OBJIndex).Name _
        & "," & MiObj.Amount _
        & "," & val _
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
Sub IniciarComercioNPC(userindex As Integer)
On Error GoTo errhandler

Call UpdateNPCInv(True, userindex, UserList(userindex).flags.TargetNpc, 0)
Call SendData(ToIndex, userindex, 0, "INITCOM")
UserList(userindex).flags.Comerciando = True

errhandler:

End Sub
Sub NPCVentaItem(userindex As Integer, ByVal i As Integer, Cantidad As Integer, NpcIndex As Integer)
On Error GoTo errhandler
Dim Infla As Long
Dim val As Long
Dim Desc As Single

If Cantidad < 1 Then Exit Sub


Infla = (Npclist(NpcIndex).Inflacion * ObjData(Npclist(NpcIndex).Invent.Object(i).OBJIndex).Valor) / 100
Desc = Descuento(userindex)

val = Fix((ObjData(Npclist(NpcIndex).Invent.Object(i).OBJIndex).Valor + Infla) / Desc)
If val = 0 Then val = 1

If UserList(userindex).Stats.GLD >= (val * Cantidad) Then
    If Npclist(UserList(userindex).flags.TargetNpc).Invent.Object(i).Amount > 0 Or Npclist(UserList(userindex).flags.TargetNpc).InvReSpawn = 0 Then
         If Cantidad > Npclist(UserList(userindex).flags.TargetNpc).Invent.Object(i).Amount And Npclist(UserList(userindex).flags.TargetNpc).InvReSpawn = 1 Then Cantidad = Npclist(UserList(userindex).flags.TargetNpc).Invent.Object(i).Amount
         Call UserCompraObj(userindex, CInt(i), UserList(userindex).flags.TargetNpc, Cantidad)
         Call SendUserORO(userindex)
    End If
Else
    Call SendData(ToIndex, userindex, 0, "2Q")
    Exit Sub
End If

errhandler:

End Sub
Sub NPCCompraItem(userindex As Integer, ByVal Item As Byte, Cantidad As Integer)
On Error GoTo errhandler

If ObjData(UserList(userindex).Invent.Object(Item).OBJIndex).Newbie = 1 Then
    Call SendData(ToIndex, userindex, 0, "6P")
    Exit Sub
End If

If ObjData(UserList(userindex).Invent.Object(Item).OBJIndex).NoSeCae = 1 Or ObjData(UserList(userindex).Invent.Object(Item).OBJIndex).Real > 0 Or ObjData(UserList(userindex).Invent.Object(Item).OBJIndex).Caos > 0 Or ObjData(UserList(userindex).Invent.Object(Item).OBJIndex).Newbie = 1 Then
    Call SendData(ToIndex, userindex, 0, "||No puedes vender este item." & FONTTYPE_WARNING)
    Exit Sub
End If

If UserList(userindex).Invent.Object(Item).Amount > 0 And UserList(userindex).Invent.Object(Item).Equipped = 0 Then
    If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Item).Amount Then Cantidad = UserList(userindex).Invent.Object(Item).Amount
    UserList(userindex).Invent.Object(Item).Amount = UserList(userindex).Invent.Object(Item).Amount - Cantidad
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + (ObjData(UserList(userindex).Invent.Object(Item).OBJIndex).Valor / 3 * Cantidad)
    If UserList(userindex).Invent.Object(Item).Amount <= 0 Then
        UserList(userindex).Invent.Object(Item).Amount = 0
        UserList(userindex).Invent.Object(Item).OBJIndex = 0
        UserList(userindex).Invent.Object(Item).Equipped = 0
    End If
    Call SubirSkill(userindex, Comerciar)
    Call UpdateUserInv(False, userindex, Item)
End If

Call SendUserORO(userindex)
Exit Sub
errhandler:

End Sub
Public Function Descuento(userindex As Integer) As Single

Descuento = CSng(Minimo(10 + (Fix((UserList(userindex).Stats.UserSkills(Comerciar) + UserList(userindex).Stats.UserAtributos(Carisma) - 10) / 10)), 20)) / 10

End Function
