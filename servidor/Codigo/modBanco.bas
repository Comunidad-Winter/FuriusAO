Attribute VB_Name = "modBanco"

Option Explicit
Sub IniciarDeposito(userindex As Integer)
On Error GoTo errhandler

Call UpdateBancoInv(True, userindex, 0)
Call SendData(ToIndex, userindex, 0, "INITBANCO")
UserList(userindex).flags.Comerciando = True

errhandler:

End Sub
Sub UpdateBancoInv(UpdateAll As Boolean, userindex As Integer, Slot As Byte, Optional ByVal TodaInfo As Boolean)
Dim i As Byte

If UpdateAll Then
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        Call EnviarBancoItem(userindex, i, UpdateAll)
    Next
Else
    Call EnviarBancoItem(userindex, Slot, TodaInfo)
End If

End Sub
Sub EnviarBancoItem(userindex As Integer, Slot As Byte, ByVal AllInfo As Boolean)
Dim MiObj As UserOBJ

MiObj = UserList(userindex).BancoInvent.Object(Slot)

If MiObj.OBJIndex Then
    If AllInfo Then
        Call SendData(ToIndex, userindex, 0, "OTII" & Slot _
        & "," & ObjData(MiObj.OBJIndex).Name _
        & "," & MiObj.Amount _
        & "," & 0 _
        & "," & ObjData(MiObj.OBJIndex).GrhIndex _
        & "," & MiObj.OBJIndex _
        & "," & ObjData(MiObj.OBJIndex).ObjType _
        & "," & ObjData(MiObj.OBJIndex).MaxHit _
        & "," & ObjData(MiObj.OBJIndex).MinHit _
        & "," & ObjData(MiObj.OBJIndex).MaxDef _
        & "," & ObjData(MiObj.OBJIndex).MinDef _
        & "," & ObjData(MiObj.OBJIndex).TipoPocion _
        & "," & ObjData(MiObj.OBJIndex).MaxModificador _
        & "," & ObjData(MiObj.OBJIndex).MinModificador)
    Else
        Call SendData(ToIndex, userindex, 0, "OTIC" & Slot & "," & MiObj.Amount)
    End If
Else
    Call SendData(ToIndex, userindex, 0, "OTIV" & Slot)
End If

End Sub
Sub UserRetiraItem(userindex As Integer, ByVal i As Byte, Cantidad As Integer)
On Error GoTo errhandler

If Cantidad < 1 Then Exit Sub

If UserList(userindex).BancoInvent.Object(i).Amount Then
     If Cantidad > UserList(userindex).BancoInvent.Object(i).Amount Then Cantidad = UserList(userindex).BancoInvent.Object(i).Amount
     Call UserReciveObj(userindex, CInt(i), Cantidad)
     Call UpdateBancoInv(False, userindex, i)
End If

errhandler:

End Sub
Sub UserReciveObj(userindex As Integer, ByVal OBJIndex As Integer, Cantidad As Integer)
Dim Slot As Byte
Dim ObjI As Integer


If UserList(userindex).BancoInvent.Object(OBJIndex).Amount <= 0 Then Exit Sub

ObjI = UserList(userindex).BancoInvent.Object(OBJIndex).OBJIndex



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
                Call SendData(ToIndex, userindex, 0, "5W")
                Exit Sub
            End If
        Loop
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If




If UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    
    UserList(userindex).Invent.Object(Slot).OBJIndex = ObjI
    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + Cantidad
    
                
    Call UpdateUserInv(False, userindex, Slot)
    Call QuitarBancoInvItem(userindex, CByte(OBJIndex), Cantidad)
    
Else
    Call SendData(ToIndex, userindex, 0, "5W")
End If


End Sub

Sub QuitarBancoInvItem(userindex As Integer, Slot As Byte, Cantidad As Integer)
Dim OBJIndex As Integer
OBJIndex = UserList(userindex).BancoInvent.Object(Slot).OBJIndex

UserList(userindex).BancoInvent.Object(Slot).Amount = UserList(userindex).BancoInvent.Object(Slot).Amount - Cantidad

If UserList(userindex).BancoInvent.Object(Slot).Amount <= 0 Then
    UserList(userindex).BancoInvent.NroItems = UserList(userindex).BancoInvent.NroItems - 1
    UserList(userindex).BancoInvent.Object(Slot).OBJIndex = 0
    UserList(userindex).BancoInvent.Object(Slot).Amount = 0
End If

End Sub
Sub UserDepositaItem(userindex As Integer, ByVal Item As Integer, Cantidad As Integer)
On Error GoTo errhandler
   
If UserList(userindex).Invent.Object(Item).Amount > 0 And UserList(userindex).Invent.Object(Item).Equipped = 0 Then
    If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Item).Amount Then Cantidad = UserList(userindex).Invent.Object(Item).Amount
    Call UserDejaObj(userindex, CInt(Item), Cantidad)
End If

errhandler:

End Sub
Sub UserDejaObj(userindex As Integer, ByVal OBJIndex As Integer, Cantidad As Integer)
Dim Slot As Byte
Dim ObjI As Integer

If Cantidad < 1 Then Exit Sub

ObjI = UserList(userindex).Invent.Object(OBJIndex).OBJIndex

Slot = 1
Do Until UserList(userindex).BancoInvent.Object(Slot).OBJIndex = ObjI And _
    UserList(userindex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    Slot = Slot + 1
    
    If Slot > MAX_BANCOINVENTORY_SLOTS Then Exit Do
Loop

If Slot > MAX_BANCOINVENTORY_SLOTS Then
    Slot = 1
    Do Until UserList(userindex).BancoInvent.Object(Slot).OBJIndex = 0
        Slot = Slot + 1
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Call SendData(ToIndex, userindex, 0, "9Y")
            Exit Sub
            Exit Do
        End If
    Loop
    If Slot <= MAX_BANCOINVENTORY_SLOTS Then UserList(userindex).BancoInvent.NroItems = UserList(userindex).BancoInvent.NroItems + 1
End If

If UserList(userindex).Tienda.NroItems + UserList(userindex).BancoInvent.NroItems > MAX_BANCOINVENTORY_SLOTS Then
    Call SendData(ToIndex, userindex, 0, "/L")
    Exit Sub
End If

If Slot <= MAX_BANCOINVENTORY_SLOTS Then
    If UserList(userindex).BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        UserList(userindex).BancoInvent.Object(Slot).OBJIndex = ObjI
        UserList(userindex).BancoInvent.Object(Slot).Amount = UserList(userindex).BancoInvent.Object(Slot).Amount + Cantidad
        Call QuitarUserInvItem(userindex, CByte(OBJIndex), Cantidad)
        Call UpdateBancoInv(False, userindex, Slot, True)
    Else
        Call SendData(ToIndex, userindex, 0, "9Y")
    End If
    Call UpdateUserInv(False, userindex, CByte(OBJIndex))
Else
    Call QuitarUserInvItem(userindex, CByte(OBJIndex), Cantidad)
End If

End Sub


