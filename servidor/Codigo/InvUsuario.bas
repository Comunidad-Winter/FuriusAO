Attribute VB_Name = "InvUsuario"
    

Option Explicit
Public Sub AcomodarItems(userindex As Integer, Item1 As Byte, Item2 As Byte)
Dim tObj As UserOBJ
Dim tObj2 As UserOBJ

tObj = UserList(userindex).Invent.Object(Item1)
tObj2 = UserList(userindex).Invent.Object(Item2)

UserList(userindex).Invent.Object(Item1) = tObj2
UserList(userindex).Invent.Object(Item2) = tObj

If tObj.Equipped = 1 Then
    Select Case ObjData(tObj.OBJIndex).ObjType
        Case OBJTYPE_WEAPON
            UserList(userindex).Invent.WeaponEqpSlot = Item2
        Case OBJTYPE_HERRAMIENTAS
            UserList(userindex).Invent.HerramientaEqpslot = Item2
        Case OBJTYPE_BARCOS
            UserList(userindex).Invent.BarcoSlot = Item2
        Case OBJTYPE_ARMOUR
            Select Case ObjData(tObj.OBJIndex).SubTipo
                Case OBJTYPE_CASCO
                    UserList(userindex).Invent.CascoEqpSlot = Item2
                Case OBJTYPE_ARMADURA
                    UserList(userindex).Invent.ArmourEqpSlot = Item2
                Case OBJTYPE_ESCUDO
                    UserList(userindex).Invent.EscudoEqpSlot = Item2
            End Select
        Case OBJTYPE_FLECHAS
            UserList(userindex).Invent.MunicionEqpSlot = Item2
    End Select
End If

If tObj2.Equipped = 1 Then
    Select Case ObjData(tObj2.OBJIndex).ObjType
        Case OBJTYPE_WEAPON
            UserList(userindex).Invent.WeaponEqpSlot = Item1
        Case OBJTYPE_HERRAMIENTAS
            UserList(userindex).Invent.HerramientaEqpslot = Item1
        Case OBJTYPE_BARCOS
            UserList(userindex).Invent.BarcoSlot = Item1
        Case OBJTYPE_ARMOUR
            Select Case ObjData(tObj2.OBJIndex).SubTipo
                Case OBJTYPE_CASCO
                    UserList(userindex).Invent.CascoEqpSlot = Item1
                Case OBJTYPE_ARMADURA
                    UserList(userindex).Invent.ArmourEqpSlot = Item1
                Case OBJTYPE_ESCUDO
                    UserList(userindex).Invent.EscudoEqpSlot = Item1
            End Select
        Case OBJTYPE_FLECHAS
            UserList(userindex).Invent.MunicionEqpSlot = Item1
    End Select
End If

Call UpdateUserInv(False, userindex, Item1)
Call UpdateUserInv(False, userindex, Item2)

End Sub

Public Sub CalcularSta(userindex As Integer)

Select Case UserList(userindex).Clase
    Case CIUDADANO, TRABAJADOR, EXPERTO_MINERALES
        UserList(userindex).Stats.MaxSta = 15 * UserList(userindex).Stats.ELV
    Case MINERO
        UserList(userindex).Stats.MaxSta = (15 + AdicionalSTMinero) * UserList(userindex).Stats.ELV
    Case HERRERO
        UserList(userindex).Stats.MaxSta = 15 * UserList(userindex).Stats.ELV
    Case TALADOR
        UserList(userindex).Stats.MaxSta = (15 + AdicionalSTLeñador) * UserList(userindex).Stats.ELV
    Case CARPINTERO
        UserList(userindex).Stats.MaxSta = 15 * UserList(userindex).Stats.ELV
    Case PESCADOR
        UserList(userindex).Stats.MaxSta = (15 + AdicionalSTPescador) * UserList(userindex).Stats.ELV
    Case Is <= 37
        UserList(userindex).Stats.MaxSta = 15 * UserList(userindex).Stats.ELV
    Case MAGO, NIGROMANTE
        UserList(userindex).Stats.MaxSta = (15 - AdicionalSTLadron / 2) * UserList(userindex).Stats.ELV
    Case Else
        UserList(userindex).Stats.MaxSta = 15 * UserList(userindex).Stats.ELV
End Select

UserList(userindex).Stats.MaxSta = 60 + UserList(userindex).Stats.MaxSta
UserList(userindex).Stats.MinSta = Minimo(UserList(userindex).Stats.MinSta, UserList(userindex).Stats.MaxSta)

End Sub
Public Sub VerObjetosEquipados(userindex As Integer)

With UserList(userindex).Invent
    If .CascoEqpSlot Then
        .Object(.CascoEqpSlot).Equipped = 1
        .CascoEqpObjIndex = .Object(.CascoEqpSlot).OBJIndex
        UserList(userindex).Char.CascoAnim = ObjData(.CascoEqpObjIndex).CascoAnim
    Else
        UserList(userindex).Char.CascoAnim = NingunCasco
    End If
    
    If .BarcoSlot Then .BarcoObjIndex = .Object(.BarcoSlot).OBJIndex
    
    If .ArmourEqpSlot Then
        .Object(.ArmourEqpSlot).Equipped = 1
        .ArmourEqpObjIndex = .Object(.ArmourEqpSlot).OBJIndex
        UserList(userindex).Char.Body = ObjData(.ArmourEqpObjIndex).Ropaje
    Else
        Call DarCuerpoDesnudo(userindex)
    End If
    
    If .WeaponEqpSlot Then
        .Object(.WeaponEqpSlot).Equipped = 1
        .WeaponEqpObjIndex = .Object(.WeaponEqpSlot).OBJIndex
        UserList(userindex).Char.WeaponAnim = ObjData(.WeaponEqpObjIndex).WeaponAnim
    End If
    
    If .EscudoEqpSlot Then
        .Object(.EscudoEqpSlot).Equipped = 1
        .EscudoEqpObjIndex = .Object(.EscudoEqpSlot).OBJIndex
        UserList(userindex).Char.ShieldAnim = ObjData(.EscudoEqpObjIndex).ShieldAnim
    Else
        UserList(userindex).Char.ShieldAnim = NingunEscudo
    End If

    If .MunicionEqpSlot Then
        .Object(.MunicionEqpSlot).Equipped = 1
        .MunicionEqpObjIndex = .Object(.MunicionEqpSlot).OBJIndex
    End If
    
    If .HerramientaEqpslot Then
        .Object(.HerramientaEqpslot).Equipped = 1
        .HerramientaEqpObjIndex = .Object(.HerramientaEqpslot).OBJIndex
    End If
End With

End Sub
Public Function TieneObjetosRobables(userindex As Integer) As Boolean
On Error Resume Next
Dim i As Byte

For i = 1 To MAX_INVENTORY_SLOTS
    If ObjEsRobable(userindex, i) Then
        TieneObjetosRobables = True
        Exit For
    End If
Next

End Function
Function ClaseBase(Clase As Byte) As Boolean

ClaseBase = (Clase = CIUDADANO Or Clase = TRABAJADOR Or Clase = EXPERTO_MINERALES Or _
            Clase = EXPERTO_MADERA Or Clase = LUCHADOR Or Clase = CON_MANA Or _
            Clase = HECHICERO Or Clase = ORDEN_SAGRADA Or Clase = NATURALISTA Or _
            Clase = SIGILOSO Or Clase = SIN_MANA Or Clase = BANDIDO Or _
            Clase = CABALLERO)

End Function
Function ClaseMana(Clase As Byte) As Boolean

ClaseMana = (Clase >= CON_MANA And Clase < SIN_MANA)

End Function
Function ClaseNoMana(Clase As Byte) As Boolean

ClaseNoMana = (Clase >= SIN_MANA)

End Function
Function ClaseTrabajadora(Clase As Byte) As Boolean

ClaseTrabajadora = (Clase > CIUDADANO And Clase < LUCHADOR)

End Function
Function ClasePuedeHechizo(userindex As Integer, ByVal OBJIndex As Integer) As Boolean
On Error GoTo manejador
Dim flag As Boolean

If UserList(userindex).flags.Privilegios > 1 Then
    ClasePuedeHechizo = True
    Exit Function
End If

'If ObjData(OBJIndex).UnicaClase <> 0 Then
'If ObjData(OBJIndex).UnicaClase = UserList(userindex).Clase Then
'ClasePuedeHechizo = False
'Exit Function
'End If
'End If


If ObjData(OBJIndex).ClaseProhibida(1) > 0 Then
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(OBJIndex).ClaseProhibida(i) = UserList(userindex).Clase Then
            ClasePuedeHechizo = True
            Exit Function
        End If
    Next
Else: ClasePuedeHechizo = True
End If

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarHechizo")
End Function
Function ClasePuedeUsarItem(userindex As Integer, ByVal OBJIndex As Integer) As Boolean
On Error GoTo manejador
Dim flag As Boolean

If UserList(userindex).flags.Privilegios Then
   ClasePuedeUsarItem = True
   Exit Function
End If

If ObjData(OBJIndex).UnicaClase <> 0 Then
If ObjData(OBJIndex).UnicaClase = UserList(userindex).Clase Then
ClasePuedeUsarItem = False
Exit Function
End If
End If


If Len(ObjData(OBJIndex).ClaseProhibida(1)) > 0 Then
    Dim i As Integer
    For i = 1 To NUMCLASES
    
        If ObjData(OBJIndex).ClaseProhibida(i) = UserList(userindex).Clase Then
            ClasePuedeUsarItem = False
            Exit Function
        ElseIf ObjData(OBJIndex).ClaseProhibida(i) = 0 Then
            Exit For
        End If
    Next
End If

If UserList(userindex).Clase = Druida Then
    If ObjData(OBJIndex).Druida > 0 Then
     If UserList(userindex).Recompensas(2) = ObjData(OBJIndex).Druida Then
    ClasePuedeUsarItem = False
    Call SendData(ToIndex, userindex, 0, "||No tienes la recompensa necesaria para utilizar este item" & FONTTYPE_INFO)
    Exit Function
     End If
    End If
End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function
Function RazaPuedeUsarItem(userindex As Integer, ByVal OBJIndex As Integer) As Boolean
On Error GoTo manejador
Dim flag As Boolean

If UserList(userindex).flags.Privilegios Then
    RazaPuedeUsarItem = True
    Exit Function
End If

        If Len(ObjData(OBJIndex).RazaProhibida(1)) > 0 Then
            Dim i As Integer
            For i = 1 To NUMRAZAS
                If (ObjData(OBJIndex).RazaProhibida(i)) = (UserList(userindex).Raza) Then
                    RazaPuedeUsarItem = False
                    Exit Function
                End If
            Next
            RazaPuedeUsarItem = True
        Else
            RazaPuedeUsarItem = True
        End If
        
Exit Function

manejador:
    LogError ("Error en RazaPuedeUsarItem")
End Function
Sub QuitarNewbieObj(userindex As Integer)
Dim j As Byte

For j = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(j).OBJIndex Then
        If ObjData(UserList(userindex).Invent.Object(j).OBJIndex).Newbie = 1 Then _
            Call QuitarVariosItem(userindex, j, MAX_INVENTORY_OBJS)
            Call UpdateUserInv(False, userindex, j)
    End If
Next

End Sub

Sub TirarOro(ByVal Cantidad As Long, userindex As Integer)
On Error GoTo errhandler
Dim nPos As WorldPos
If Cantidad > 100000 Then Exit Sub

If Cantidad <= 999 Or Cantidad > UserList(userindex).Stats.GLD Then Exit Sub

Dim MiObj As Obj

MiObj.OBJIndex = iORO

If UserList(userindex).flags.Privilegios Then Call LogGM(UserList(userindex).Name, "Tiro cantidad:" & Cantidad & " Objeto:" & ObjData(MiObj.OBJIndex).Name, False)

Do While Cantidad > 0
    MiObj.Amount = Minimo(Cantidad, MAX_INVENTORY_OBJS)
        
    nPos = TirarItemAlPiso(UserList(userindex).POS, MiObj)
    If nPos.Map = 0 Then Exit Sub
    
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - MiObj.Amount
    Cantidad = Cantidad - MiObj.Amount
Loop
    
Exit Sub

errhandler:

End Sub
Sub QuitarUserInvItem(userindex As Integer, ByVal Slot As Byte, Cantidad As Integer)
Dim MiObj As Obj

If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)

UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - Cantidad

If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).OBJIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
End If
    
End Sub
Sub QuitarUnItem(userindex As Integer, ByVal Slot As Byte)
Dim MiObj As Obj

If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 And UserList(userindex).Invent.Object(Slot).Amount = 1 Then Call Desequipar(userindex, Slot)

UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - 1

If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).OBJIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
    Call SendData(ToIndex, userindex, 0, "3I" & Slot)
Else
    Call SendData(ToIndex, userindex, 0, "2I" & Slot)
End If

End Sub
Sub QuitarBebida(userindex As Integer, ByVal Slot As Byte)

Dim MiObj As Obj

If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)


    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - 1


If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).OBJIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
    Call SendData(ToIndex, userindex, 0, "6I" & Slot & "," & UserList(userindex).Stats.MinAGU)
    Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "46")
    
Else
Call SendData(ToIndex, userindex, 0, "6J" & Slot & "," & UserList(userindex).Stats.MinAGU)
Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "46")

End If
    
End Sub
Sub QuitarComida(userindex As Integer, ByVal Slot As Byte)

Dim MiObj As Obj

If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)


    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - 1


If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).OBJIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
    Call SendData(ToIndex, userindex, 0, "7K" & Slot & "," & UserList(userindex).Stats.MinHam)
    Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "7")

Else
Call SendData(ToIndex, userindex, 0, "6K" & Slot & "," & UserList(userindex).Stats.MinHam)
Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "7")

End If
    
End Sub

Sub QuitarPocion(userindex As Integer, ByVal Slot As Byte)

Dim MiObj As Obj

If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)


    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - 1

If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).OBJIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
    Call SendData(ToIndex, userindex, 0, "4J" & Slot)
    Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "46")

Else
Call SendData(ToIndex, userindex, 0, "3J" & Slot)
Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "46")
End If
    
End Sub

Sub QuitarPocionMana(userindex As Integer, ByVal Slot As Byte)

Dim MiObj As Obj

If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)


UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - 1


If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).OBJIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
    Call SendData(ToIndex, userindex, 0, "8I" & Slot & "," & UserList(userindex).Stats.MinMAN)
    Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "46")

Else
Call SendData(ToIndex, userindex, 0, "7I" & Slot & "," & UserList(userindex).Stats.MinMAN)
Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "46")

End If
    
End Sub
Sub QuitarPocionVida(userindex As Integer, ByVal Slot As Byte)

Dim MiObj As Obj

If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)


    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - 1

If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).OBJIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
    Call SendData(ToIndex, userindex, 0, "2J" & Slot & "," & UserList(userindex).Stats.MinHP)
    Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "46")

Else
Call SendData(ToIndex, userindex, 0, "9I" & Slot & "," & UserList(userindex).Stats.MinHP)
Call SendData(ToPCAreaButIndex, userindex, UserList(userindex).POS.Map, "TW" & "46")

End If
    
End Sub
Sub QuitarVariosItem(userindex As Integer, ByVal Slot As Byte, Cantidad As Integer)

Dim MiObj As Obj

If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 And UserList(userindex).Invent.Object(Slot).Amount <= Cantidad Then Call Desequipar(userindex, Slot)


UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - Cantidad


If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).OBJIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
    Call SendData(ToIndex, userindex, 0, "3I" & Slot)
Else
    Call SendData(ToIndex, userindex, 0, "4I" & Slot & "," & Cantidad)
End If
    
End Sub
Sub UpdateUserInv(ByVal UpdateAll As Boolean, userindex As Integer, Slot As Byte, Optional JustAmount As Boolean)
Dim i As Byte

If UpdateAll Then
    For i = 1 To MAX_INVENTORY_SLOTS
        Call SendUserItem(userindex, i, JustAmount)
    Next
Else
    Call SendUserItem(userindex, Slot, JustAmount)
End If

End Sub
Sub DropObj(userindex As Integer, Slot As Byte, ByVal Num As Integer, Map As Integer, x As Integer, Y As Integer)
Dim Obj As Obj

If Num Then
  If Num > UserList(userindex).Invent.Object(Slot).Amount Then Num = UserList(userindex).Invent.Object(Slot).Amount
  
  
  If MapData(UserList(userindex).POS.Map, x, Y).OBJInfo.OBJIndex = 0 Then
        If UserList(userindex).Invent.Object(Slot).Equipped = 1 And UserList(userindex).Invent.Object(Slot).Amount <= Num Then Call Desequipar(userindex, Slot)
        Obj.OBJIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
        If UserList(userindex).flags.Privilegios < 2 Then
            If ObjData(Obj.OBJIndex).NoComerciable = 1 Then
                Call SendData(ToIndex, userindex, 0, "2W")
                Exit Sub
            End If
            
            If ObjData(Obj.OBJIndex).NoSeCae Then
                Call SendData(ToIndex, userindex, 0, "||No puedes tirar este objeto." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If ObjData(Obj.OBJIndex).Newbie = 1 And EsNewbie(userindex) Then
                Call SendData(ToIndex, userindex, 0, "3W")
                Exit Sub
            End If
        End If
        
        Obj.Amount = Num
        
        Call MakeObj(ToMap, 0, Map, Obj, Map, x, Y)
        Call QuitarVariosItem(userindex, Slot, Num)
        
        If UserList(userindex).flags.Privilegios Then Call LogGM(UserList(userindex).Name, "Tiro cantidad:" & Num & " Objeto:" & ObjData(Obj.OBJIndex).Name, False)
  Else
        Call SendData(ToIndex, userindex, 0, "4W")
  End If
    
End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal Num As Integer, Map As Integer, x As Integer, Y As Integer)

MapData(Map, x, Y).OBJInfo.Amount = MapData(Map, x, Y).OBJInfo.Amount - Num

If MapData(Map, x, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, x, Y).OBJInfo.OBJIndex = 0
    MapData(Map, x, Y).OBJInfo.Amount = 0
    Call SendData(sndRoute, sndIndex, sndMap, "BO" & x & "," & Y)
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, x As Integer, Y As Integer)


MapData(Map, x, Y).OBJInfo = Obj
Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.OBJIndex).GrhIndex & "," & x & "," & Y)

End Sub

Function MeterItemEnInventario(userindex As Integer, MiObj As Obj) As Boolean
On Error GoTo errhandler


 
Dim x As Integer
Dim Y As Integer
Dim Slot As Byte


Slot = 1
Do Until UserList(userindex).Invent.Object(Slot).OBJIndex = MiObj.OBJIndex And _
         UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then Exit Do
Loop
    

If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(userindex).Invent.Object(Slot).OBJIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call SendData(ToIndex, userindex, 0, "5W")
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If
    

If UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   
   UserList(userindex).Invent.Object(Slot).OBJIndex = MiObj.OBJIndex
   UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(userindex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, userindex, Slot)


Exit Function
errhandler:

End Function


Sub GetObj(userindex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj


If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).OBJInfo.OBJIndex Then
    
    If ObjData(MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).OBJInfo.OBJIndex).Agarrable <> 1 Then
        Dim x As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        x = UserList(userindex).POS.x
        Y = UserList(userindex).POS.Y
        Obj = ObjData(MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).OBJInfo.OBJIndex)
        MiObj.Amount = MapData(UserList(userindex).POS.Map, x, Y).OBJInfo.Amount
        MiObj.OBJIndex = MapData(UserList(userindex).POS.Map, x, Y).OBJInfo.OBJIndex

        If ObjData(MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_GUITA Then
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + MapData(UserList(userindex).POS.Map, x, Y).OBJInfo.Amount
        Call SendUserORO(userindex)
        Call EraseObj(ToMap, 0, UserList(userindex).POS.Map, MapData(UserList(userindex).POS.Map, x, Y).OBJInfo.Amount, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
        If UserList(userindex).flags.Privilegios Then Call LogGM(UserList(userindex).Name, "Agarro oro:" & MiObj.Amount, False)
        Exit Sub
        End If


        Dim Slotx As Byte

        
        If Not MeterItemEnInventario(userindex, MiObj) Then
        
        Else
            Call EraseObj(ToMap, 0, UserList(userindex).POS.Map, MapData(UserList(userindex).POS.Map, x, Y).OBJInfo.Amount, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
            If UserList(userindex).flags.Privilegios Then Call LogGM(UserList(userindex).Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.OBJIndex).Name, False)
        
        
        If MiObj.OBJIndex = BANDERAINDEXCIUDA Then
                Slotx = 1
                Do Until UserList(userindex).Invent.Object(Slotx).OBJIndex = BANDERAINDEXCIUDA
                Slotx = Slotx + 1
                Loop

            If UserList(userindex).Faccion.Bando = Real Then
                Call DevolverBandera(1, UserList(userindex).Name)
                Call QuitarVariosItem(userindex, Slotx, 1)
                Call UpdateUserInv(False, userindex, Slotx)
            ElseIf UserList(userindex).Faccion.Bando = Caos Then
            
                If RazaBaja(userindex) Then
                ObjData(UserList(userindex).Invent.Object(Slotx).OBJIndex).Ropaje = 325
                Else
                ObjData(UserList(userindex).Invent.Object(Slotx).OBJIndex).Ropaje = 323
                End If
                
                Call EquiparInvItem(userindex, Slotx)
                Call SendData(ToMap, 0, MAP_CTF, "||Capture the Flag> " & UserList(userindex).Name & " agarró la bandera ciudadana." & FONTTYPE_CAOS)
                Call SendData(ToMap, 0, MAP_CTC, "||Capture the Flag> " & UserList(userindex).Name & " agarró la bandera ciudadana." & FONTTYPE_CAOS)
                Call SumaPuntos(userindex, 10)
            End If
            
        ElseIf MiObj.OBJIndex = BANDERAINDEXCRIMI Then
                Slotx = 1
                Do Until UserList(userindex).Invent.Object(Slotx).OBJIndex = BANDERAINDEXCRIMI
                Slotx = Slotx + 1
                Loop
            
            If UserList(userindex).Faccion.Bando = Caos Then
                Call DevolverBandera(2, UserList(userindex).Name)
                Call QuitarVariosItem(userindex, Slotx, 1)
                Call UpdateUserInv(False, userindex, Slotx)
            ElseIf UserList(userindex).Faccion.Bando = Real Then
            
                If RazaBaja(userindex) Then
                ObjData(UserList(userindex).Invent.Object(Slotx).OBJIndex).Ropaje = 328
                Else
                ObjData(UserList(userindex).Invent.Object(Slotx).OBJIndex).Ropaje = 326
                End If
                
                Call EquiparInvItem(userindex, Slotx)
                Call SendData(ToMap, 0, MAP_CTF, "||Capture the Flag> " & UserList(userindex).Name & " agarró la bandera criminal." & FONTTYPE_ARMADA)
                Call SendData(ToMap, 0, MAP_CTC, "||Capture the Flag> " & UserList(userindex).Name & " agarró la bandera criminal." & FONTTYPE_ARMADA)
                Call SumaPuntos(userindex, 10)
            End If
        End If
        
        
        
        
        End If
        
    End If
Else
    Call SendData(ToIndex, userindex, 0, "8K")
    
End If

End Sub
Sub Desequipar(userindex As Integer, ByVal Slot As Byte)



Dim Obj As ObjData
If Slot = 0 Then Exit Sub
If UserList(userindex).Invent.Object(Slot).OBJIndex = 0 Then Exit Sub

Obj = ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex)

Select Case Obj.ObjType


    Case OBJTYPE_WEAPON

        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.WeaponEqpObjIndex = 0
        UserList(userindex).Invent.WeaponEqpSlot = 0

        Call ChangeUserArma(ToMap, 0, UserList(userindex).POS.Map, userindex, NingunArma)
        
    Case OBJTYPE_FLECHAS
    
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.MunicionEqpObjIndex = 0
        UserList(userindex).Invent.MunicionEqpSlot = 0
        
    Case OBJTYPE_HERRAMIENTAS
            
        If UserList(userindex).flags.Trabajando Then
            If UserList(userindex).flags.CodigoTrabajo Then
                Exit Sub
            Else
                Call SacarModoTrabajo(userindex)
            End If
        End If
        
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.HerramientaEqpObjIndex = 0
        UserList(userindex).Invent.HerramientaEqpslot = 0
        
    Case OBJTYPE_ARMOUR
        If UserList(userindex).flags.Montado = 1 Then Exit Sub

        Select Case Obj.SubTipo
        
            Case OBJTYPE_ARMADURA
                UserList(userindex).Invent.Object(Slot).Equipped = 0
                UserList(userindex).Invent.ArmourEqpObjIndex = 0
                UserList(userindex).Invent.ArmourEqpSlot = 0
                If UserList(userindex).flags.Transformado = 0 Then
                    Call DarCuerpoDesnudo(userindex)
                    Call ChangeUserBody(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body)
                End If
                
            Case OBJTYPE_CASCO
                UserList(userindex).Invent.Object(Slot).Equipped = 0
                UserList(userindex).Invent.CascoEqpObjIndex = 0
                UserList(userindex).Invent.CascoEqpSlot = 0
                If UserList(userindex).flags.Transformado = 0 Then
                    Call ChangeUserCasco(ToMap, 0, UserList(userindex).POS.Map, userindex, NingunCasco)
                End If
            Case OBJTYPE_ESCUDO
                UserList(userindex).Invent.Object(Slot).Equipped = 0
                UserList(userindex).Invent.EscudoEqpObjIndex = 0
                UserList(userindex).Invent.EscudoEqpSlot = 0
                If UserList(userindex).flags.Transformado = 0 Then
                    Call ChangeUserEscudo(ToMap, 0, UserList(userindex).POS.Map, userindex, NingunEscudo)
                End If
        End Select
    
End Select

Call DesequiparItem(userindex, Slot)

End Sub
Function SexoPuedeUsarItem(userindex As Integer, ByVal OBJIndex As Integer) As Boolean
On Error GoTo errhandler

If UserList(userindex).flags.Privilegios Then
    SexoPuedeUsarItem = True
    Exit Function
End If

If ObjData(OBJIndex).MUJER = 1 Then
    SexoPuedeUsarItem = UserList(userindex).Genero = MUJER
ElseIf ObjData(OBJIndex).HOMBRE = 1 Then
    SexoPuedeUsarItem = UserList(userindex).Genero = HOMBRE
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function
Function FaccionClasePuedeUsarItem(userindex As Integer, ByVal OBJIndex As Integer) As Boolean
Dim i As Integer

If UserList(userindex).flags.Privilegios Then
    FaccionClasePuedeUsarItem = True
    Exit Function
End If

For i = 1 To Minimo(UserList(userindex).Faccion.Jerarquia, 3)
    If Armaduras(UserList(userindex).Faccion.Bando, i, TipoClase(userindex), TipoRaza(userindex)) = OBJIndex Then
        FaccionClasePuedeUsarItem = True
        Exit Function
    End If
Next

End Function
Function FaccionPuedeUsarItem(userindex As Integer, ByVal OBJIndex As Integer) As Boolean

If UserList(userindex).flags.Privilegios Then
    FaccionPuedeUsarItem = True
    Exit Function
End If

If ObjData(OBJIndex).Real >= 1 Then
    FaccionPuedeUsarItem = (UserList(userindex).Faccion.Bando = Real And UserList(userindex).Faccion.Jerarquia >= ObjData(OBJIndex).Jerarquia)
ElseIf ObjData(OBJIndex).Caos >= 1 Then
    FaccionPuedeUsarItem = (UserList(userindex).Faccion.Bando = Caos And UserList(userindex).Faccion.Jerarquia >= ObjData(OBJIndex).Jerarquia)
Else: FaccionPuedeUsarItem = True
End If

End Function
Function PuedeUsarObjeto(userindex As Integer, ByVal OBJIndex As Integer) As Byte

Select Case ObjData(OBJIndex).ObjType
    Case OBJTYPE_WEAPON
    
        If Not (OBJIndex = 367 And UserList(userindex).Clase = ASESINO) Then
            If Not RazaPuedeUsarItem(userindex, OBJIndex) Then
                PuedeUsarObjeto = 5
                Exit Function
            End If
        
            If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                 PuedeUsarObjeto = 2
                 Exit Function
            End If
        End If
        
        If Not SkillPuedeUsarItem(userindex, OBJIndex) Then
            PuedeUsarObjeto = 4
            Exit Function
        End If
       
    Case OBJTYPE_HERRAMIENTAS
    
        If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
             PuedeUsarObjeto = 2
             Exit Function
        End If

    Case OBJTYPE_ARMOUR
         
         Select Case ObjData(OBJIndex).SubTipo
        
            Case OBJTYPE_ARMADURA
            
                If Not RazaPuedeUsarItem(userindex, OBJIndex) Then
                    PuedeUsarObjeto = 5
                    Exit Function
                End If
                
                If Not SexoPuedeUsarItem(userindex, OBJIndex) Then
                    PuedeUsarObjeto = 1
                    Exit Function
                End If
                 
                If ObjData(OBJIndex).Real = 0 And ObjData(OBJIndex).Caos = 0 Then
                    If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                         PuedeUsarObjeto = 2
                         Exit Function
                    End If
                Else
                    If Not FaccionPuedeUsarItem(userindex, OBJIndex) Then
                        PuedeUsarObjeto = 3
                        Exit Function
                    End If
                    If Not FaccionClasePuedeUsarItem(userindex, OBJIndex) Then
                         PuedeUsarObjeto = 2
                         Exit Function
                    End If
                End If
            
                If Not SkillPuedeUsarItem(userindex, OBJIndex) Then
                    PuedeUsarObjeto = 4
                    Exit Function
                End If

            Case OBJTYPE_CASCO
            
                 If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                      PuedeUsarObjeto = 2
                      Exit Function
                 End If
                
                 If Not SkillPuedeUsarItem(userindex, OBJIndex) Then
                     PuedeUsarObjeto = 4
                     Exit Function
                 End If
                
            Case OBJTYPE_ESCUDO
            
                If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                    PuedeUsarObjeto = 2
                    Exit Function
                End If

                 If Not SkillPuedeUsarItem(userindex, OBJIndex) Then
                     PuedeUsarObjeto = 4
                     Exit Function
                 End If
            
            Case OBJTYPE_PERGAMINOS
                If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                    PuedeUsarObjeto = 2
                    Exit Function
                End If
            
        End Select
End Select

PuedeUsarObjeto = 0

End Function
Function SkillPuedeUsarItem(userindex As Integer, ByVal OBJIndex As Integer) As Boolean

If UserList(userindex).flags.Privilegios Then
    SkillPuedeUsarItem = True
    Exit Function
End If

If ObjData(OBJIndex).SkillCombate > UserList(userindex).Stats.UserSkills(Armas) Then Exit Function
If ObjData(OBJIndex).SkillApuñalar > UserList(userindex).Stats.UserSkills(Apuñalar) Then Exit Function
If ObjData(OBJIndex).SkillProyectiles > UserList(userindex).Stats.UserSkills(Proyectiles) Then Exit Function
If ObjData(OBJIndex).SkResistencia > UserList(userindex).Stats.UserSkills(Resis) Then Exit Function
If ObjData(OBJIndex).SkDefensa > UserList(userindex).Stats.UserSkills(Defensa) Then Exit Function
If ObjData(OBJIndex).SkillTacticas > UserList(userindex).Stats.UserSkills(Tacticas) Then Exit Function

SkillPuedeUsarItem = True

End Function
Sub EquiparInvItem(userindex As Integer, Slot As Byte)
On Error GoTo errhandler


Dim Obj As ObjData
Dim OBJIndex As Integer

OBJIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
Obj = ObjData(OBJIndex)

If Obj.Newbie = 1 And Not EsNewbie(userindex) Then
     Call SendData(ToIndex, userindex, 0, "6W")
     Exit Sub
End If

If UserList(userindex).Stats.ELV < Obj.Minlvl Then
    Call SendData(ToIndex, userindex, 0, "||Nesesitas ser nivel " & Obj.Minlvl & " para usar este objeto." & FONTTYPE_INFO)
    Exit Sub
End If

Select Case Obj.ObjType
    Case OBJTYPE_QUEST

    Case OBJTYPE_WEAPON
    
        If Not (OBJIndex = 367 And UserList(userindex).Clase = ASESINO) Then
            If Not RazaPuedeUsarItem(userindex, OBJIndex) Then
                Call SendData(ToIndex, userindex, 0, "8W")
                Exit Sub
            End If
        
            If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                 Call SendData(ToIndex, userindex, 0, "2X")
                 Exit Sub
            End If
        End If
        
        If Not SkillPuedeUsarItem(userindex, OBJIndex) Then
            Call SendData(ToIndex, userindex, 0, "7W")
            Exit Sub
        End If
                  
            If UserList(userindex).Invent.Object(Slot).Equipped Then
                Call Desequipar(userindex, Slot)
                Exit Sub
            End If
            
            
            If UserList(userindex).Invent.WeaponEqpObjIndex Then Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)

            UserList(userindex).Invent.Object(Slot).Equipped = 1
            UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
            UserList(userindex).Invent.WeaponEqpSlot = Slot
            
            
            If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil > 0 And UserList(userindex).Invent.EscudoEqpSlot > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
                Call ChangeUserEscudo(ToMap, 0, UserList(userindex).POS.Map, userindex, 0)
           End If
            
            Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SOUND_SACARARMA)

            Call ChangeUserArma(ToMap, 0, UserList(userindex).POS.Map, userindex, Obj.WeaponAnim)
            Call EquiparItem(userindex, Slot)
       
    Case OBJTYPE_HERRAMIENTAS
        If Not RazaPuedeUsarItem(userindex, OBJIndex) Then
            Call SendData(ToIndex, userindex, 0, "8W")
            Exit Sub
        End If
    
    
        If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
             Call SendData(ToIndex, userindex, 0, "2X")
             Exit Sub
        End If
       
        If OBJIndex = 753 And Not (UserList(userindex).Clase = MINERO And UserList(userindex).Recompensas(2) = 2) Then
            Call SendData(ToIndex, userindex, 0, "||Debes tener la recompensa Pica Fuerte para usar este item." & FONTTYPE_BLANCO)
            Exit Sub
        End If
        
        If UserList(userindex).Invent.Object(Slot).Equipped Then
            
            Call Desequipar(userindex, Slot)
            Exit Sub
        End If
        
        
        If UserList(userindex).Invent.HerramientaEqpObjIndex Then
            Call Desequipar(userindex, UserList(userindex).Invent.HerramientaEqpslot)
        End If

        UserList(userindex).Invent.Object(Slot).Equipped = 1
        UserList(userindex).Invent.HerramientaEqpObjIndex = OBJIndex
        UserList(userindex).Invent.HerramientaEqpslot = Slot
        Call EquiparItem(userindex, Slot)
                
    Case OBJTYPE_FLECHAS
        
         
         If UserList(userindex).Invent.Object(Slot).Equipped Then
             
             Call Desequipar(userindex, Slot)
             Exit Sub
         End If
         
         
         If UserList(userindex).Invent.MunicionEqpObjIndex Then
             Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
         End If
 
         UserList(userindex).Invent.Object(Slot).Equipped = 1
         UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
         UserList(userindex).Invent.MunicionEqpSlot = Slot
         Call EquiparItem(userindex, Slot)
    
    Case OBJTYPE_ARMOUR

         If UserList(userindex).flags.Navegando = 1 Then Exit Sub
         
         Select Case Obj.SubTipo
         
            Case OBJTYPE_ARMADURA
            
                If Not RazaPuedeUsarItem(userindex, OBJIndex) Then
                    Call SendData(ToIndex, userindex, 0, "8W")
                    Exit Sub
                End If
                
                If Not SexoPuedeUsarItem(userindex, OBJIndex) Then
                    Call SendData(ToIndex, userindex, 0, "8W")
                    Exit Sub
                End If
                 
                If ObjData(OBJIndex).Real = 0 And ObjData(OBJIndex).Caos = 0 Then
                    If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                        Call SendData(ToIndex, userindex, 0, "2X")
                        Exit Sub
                    End If
                Else
                    If Not FaccionPuedeUsarItem(userindex, OBJIndex) Then
                        Call SendData(ToIndex, userindex, 0, "%?")
                        Exit Sub
                    End If
                    If Not FaccionClasePuedeUsarItem(userindex, OBJIndex) Then
                        Call SendData(ToIndex, userindex, 0, "||Tu clase o raza no puede usar ese objeto." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                
                If Not SkillPuedeUsarItem(userindex, OBJIndex) Then
                    Call SendData(ToIndex, userindex, 0, "7W")
                    Exit Sub
                End If
                   
               
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    If UserList(userindex).Invent.Object(Slot).OBJIndex = BANDERAINDEXCRIMI Or UserList(userindex).Invent.Object(Slot).OBJIndex = BANDERAINDEXCIUDA Then Exit Sub  ' CAPTURE THE FLAG
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
        
                
                If UserList(userindex).Invent.ArmourEqpObjIndex Then
                If UserList(userindex).Invent.ArmourEqpObjIndex = BANDERAINDEXCRIMI Or UserList(userindex).Invent.ArmourEqpObjIndex = BANDERAINDEXCIUDA Then Exit Sub ' CAPTURE THE FLAG
                    Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
                End If
        
                
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
                UserList(userindex).Invent.ArmourEqpSlot = Slot
                    
                UserList(userindex).flags.Desnudo = 0
                    
                If UserList(userindex).flags.Transformado = 0 Then Call ChangeUserBody(ToMap, 0, UserList(userindex).POS.Map, userindex, Obj.Ropaje)
                Call EquiparItem(userindex, Slot)

            Case OBJTYPE_CASCO
            
            
                    If OBJIndex = 572 Then
                        If UserList(userindex).Raza <> GNOMO Then
                        Call SendData(ToIndex, userindex, 0, "||Tu Raza no puede usar este Item." & FONTTYPE_INFO)
                        Exit Sub
                        End If
                    End If

                 If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                      Call SendData(ToIndex, userindex, 0, "2X")
                      Exit Sub
                 End If
                
                 If Not SkillPuedeUsarItem(userindex, OBJIndex) Then
                     Call SendData(ToIndex, userindex, 0, "7W")
                     Exit Sub
                 End If
                 
        
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
        
                
                If UserList(userindex).Invent.CascoEqpObjIndex Then
                    Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
                End If
        
                
                
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.CascoEqpObjIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
                UserList(userindex).Invent.CascoEqpSlot = Slot
            
                Call ChangeUserCasco(ToMap, 0, UserList(userindex).POS.Map, userindex, Obj.CascoAnim)
                Call EquiparItem(userindex, Slot)
                
            Case OBJTYPE_ESCUDO
            
                If Not ClasePuedeUsarItem(userindex, OBJIndex) Then
                    Call SendData(ToIndex, userindex, 0, "2X")
                    Exit Sub
                End If
                
                If Not SkillPuedeUsarItem(userindex, OBJIndex) Then
                    Call SendData(ToIndex, userindex, 0, "7W")
                    Exit Sub
                End If
                
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
                
                
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
        
                
                If UserList(userindex).Invent.EscudoEqpObjIndex Then Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
        
                
                
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.EscudoEqpObjIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
                UserList(userindex).Invent.EscudoEqpSlot = Slot
               'FuriusAO El arquero usa escudo de tortuga
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil Then
                 '   Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                    Call ChangeUserArma(ToMap, 0, UserList(userindex).POS.Map, userindex, 0)
                End If
           ' el arquero ahora usa escudo de tortuga
                Call ChangeUserEscudo(ToMap, 0, UserList(userindex).POS.Map, userindex, Obj.ShieldAnim)
                Call EquiparItem(userindex, Slot)

        End Select
End Select


Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(userindex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler


If UserList(userindex).Raza = HUMANO Or _
   UserList(userindex).Raza = ELFO Or _
   UserList(userindex).Raza = ELFO_OSCURO Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If


Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function
Sub SacarModoTrabajo(userindex As Integer)

UserList(userindex).flags.Trabajando = 0
UserList(userindex).TrabajoPos.x = 0
UserList(userindex).TrabajoPos.Y = 0
UserList(userindex).flags.CodigoTrabajo = 0

Call SendData(ToIndex, userindex, 0, "%I")
Call SendData(ToIndex, userindex, 0, "MT")

End Sub
Sub UseInvItem(userindex As Integer, Slot As Byte, ByVal Click As Byte)
Dim Obj As ObjData
Dim OBJIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

Obj = ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex)

If Obj.Newbie = 1 And Not EsNewbie(userindex) Then
    Call SendData(ToIndex, userindex, 0, "6W")
    Exit Sub
End If

OBJIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
UserList(userindex).flags.TargetObjInvIndex = OBJIndex
UserList(userindex).flags.TargetObjInvslot = Slot

Select Case Obj.ObjType
    Case OBJTYPE_MASCOTA
    If UserList(userindex).flags.Montado Then
    UserList(userindex).Invent.MascotaEqpObjIndex = 0
    UserList(userindex).Invent.MascotaEqpSlot = 0
    UserList(userindex).flags.Montado = 0
    UserList(userindex).Char = UserList(userindex).OrigChar
    
            
            Else
    
    
    If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
    End If
    UserList(userindex).OrigChar = UserList(userindex).Char
    UserList(userindex).flags.Montado = 1
    UserList(userindex).Invent.MascotaEqpObjIndex = OBJIndex
    UserList(userindex).Invent.MascotaEqpSlot = Slot
    UserList(userindex).Char.Body = ObjData(OBJIndex).Ropaje
    UserList(userindex).Char.ShieldAnim = 0
    UserList(userindex).Char.WeaponAnim = 0
    End If
    Call ChangeUserCharB(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
    Exit Sub
    Case OBJTYPE_MERCENARIO
    If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
    End If
    
    If UserList(userindex).NroMascotas > 2 Then
    Exit Sub
    End If
    
    Call InvocacionMercenario(userindex, Obj.InvocaN)
    Call SendData(ToIndex, userindex, 0, "||Has invocado un mercenario" & FONTTYPE_VENENO)
    Call QuitarPocion(userindex, Slot)
    
    Case OBJTYPE_USEONCE
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If

        Call AddtoVar(UserList(userindex).Stats.MinHam, Obj.MinHam, UserList(userindex).Stats.MaxHam)
        UserList(userindex).flags.Hambre = 0
        
        Call QuitarComida(userindex, Slot)
            
    Case OBJTYPE_GUITA
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        
        UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + UserList(userindex).Invent.Object(Slot).Amount
        UserList(userindex).Invent.Object(Slot).Amount = 0
        UserList(userindex).Invent.Object(Slot).OBJIndex = 0
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
        Call SendUserORO(userindex)
        
    Case OBJTYPE_WEAPON
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If

        If ObjData(OBJIndex).proyectil = 1 Then
            If TiempoTranscurrido(UserList(userindex).Counters.LastFlecha) < IntervaloUserFlechas Then Exit Sub
            If TiempoTranscurrido(UserList(userindex).Counters.LastHechizo) < IntervaloUserPuedeHechiGolpe Then Exit Sub
            Call SendData(ToIndex, userindex, 0, "T01" & Proyectiles)
        Else
            If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
            If ObjData(UserList(userindex).flags.TargetObj).ObjType = OBJTYPE_LEÑA And UserList(userindex).Invent.Object(Slot).OBJIndex = DAGA Then Call TratarDeHacerFogata(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY, userindex)
        End If
        
    Case OBJTYPE_POCIONES
        If TiempoTranscurrido(UserList(userindex).Counters.LastGolpe) < (IntervaloUserPuedeAtacar / 2) Then
            Call SendData(ToIndex, userindex, 0, "6X")
            Exit Sub
        End If
                
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
                
       

        Select Case Obj.TipoPocion
        
            Case 1
                UserList(userindex).flags.DuracionEfecto = Timer
                UserList(userindex).flags.TomoPocion = True
                
                Call AddtoVar(UserList(userindex).Stats.UserAtributos(Agilidad), RandomNumber(Obj.MinModificador, Obj.MaxModificador), Minimo(UserList(userindex).Stats.UserAtributosBackUP(Agilidad) * 2, MAXATRIBUTOS))
                Call UpdateFuerzaYAg(userindex)
                
                Call QuitarPocion(userindex, Slot)
                
        
            Case 2
                UserList(userindex).flags.DuracionEfecto = Timer
                UserList(userindex).flags.TomoPocion = True
                
                Call AddtoVar(UserList(userindex).Stats.UserAtributos(fuerza), RandomNumber(Obj.MinModificador, Obj.MaxModificador), Minimo(UserList(userindex).Stats.UserAtributosBackUP(fuerza) * 2, MAXATRIBUTOS))
                Call UpdateFuerzaYAg(userindex)
                
                Call QuitarPocion(userindex, Slot)
                
            Case 3
                
                AddtoVar UserList(userindex).Stats.MinHP, RandomNumber(Obj.MinModificador, Obj.MaxModificador), UserList(userindex).Stats.MaxHP
                
                
                Call QuitarPocionVida(userindex, Slot)
                
               
               
               
               
            
            Case 4
                
                Call AddtoVar(UserList(userindex).Stats.MinMAN, Porcentaje(UserList(userindex).Stats.MaxMAN, Obj.MaxModificador), UserList(userindex).Stats.MaxMAN)
                
                
                Call QuitarPocionMana(userindex, Slot)
            Case 5
                If UserList(userindex).flags.Envenenado = 1 Then
                    UserList(userindex).flags.Envenenado = 0
                    Call SendData(ToIndex, userindex, 0, "8X")
                End If
                
                Call QuitarPocion(userindex, Slot)
                   
       End Select
       
     Case OBJTYPE_BEBIDA
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        AddtoVar UserList(userindex).Stats.MinAGU, Obj.MinSed, UserList(userindex).Stats.MaxAGU
        UserList(userindex).flags.Sed = 0
        
        
        Call QuitarBebida(userindex, Slot)
    
    Case OBJTYPE_LLAVES
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        
        If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(userindex).flags.TargetObj)
        
        If TargObj.ObjType = OBJTYPE_PUERTAS Then
            
            If TargObj.Cerrada = 1 Then
                  
                  If TargObj.Llave Then
                     If TargObj.Clave = Obj.Clave Then
         
                        MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.OBJIndex _
                        = ObjData(MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.OBJIndex).IndexCerrada
                        UserList(userindex).flags.TargetObj = MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.OBJIndex
                        Call SendData(ToIndex, userindex, 0, "9X")
                        Exit Sub
                     Else
                        Call SendData(ToIndex, userindex, 0, "2Y")
                        Exit Sub
                     End If
                  Else
                     If TargObj.Clave = Obj.Clave Then
                        MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.OBJIndex _
                        = ObjData(MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.OBJIndex).IndexCerradaLlave
                        Call SendData(ToIndex, userindex, 0, "3Y")
                        UserList(userindex).flags.TargetObj = MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.OBJIndex
                        Exit Sub
                     Else
                        Call SendData(ToIndex, userindex, 0, "2Y")
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(ToIndex, userindex, 0, "4Y")
                  Exit Sub
            End If
            
        End If
    
    Case OBJTYPE_BOTELLAVACIA
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        If MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).Agua = 0 Then
            Call SendData(ToIndex, userindex, 0, "9F")
            Exit Sub
        End If
        MiObj.Amount = 1
        MiObj.OBJIndex = ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex).IndexAbierta
        If Not MeterItemEnInventario(userindex, MiObj) Then Exit Sub
        '    Call TirarItemAlPiso(UserList(userindex).POS, MiObj)
        'End If
        Call QuitarUnItem(userindex, Slot)

            
    Case OBJTYPE_BOTELLALLENA
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        AddtoVar UserList(userindex).Stats.MinAGU, Obj.MinSed, UserList(userindex).Stats.MaxAGU
        UserList(userindex).flags.Sed = 0
        Call EnviarHyS(userindex)
        MiObj.Amount = 1
        MiObj.OBJIndex = ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex).IndexCerrada
        Call QuitarUnItem(userindex, Slot)
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(userindex).POS, MiObj)
        End If
             
    Case OBJTYPE_HERRAMIENTAS

        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        
        If UserList(userindex).Stats.MinSta = 0 Then
            Call SendData(ToIndex, userindex, 0, "9E")
            Exit Sub
        End If

        If UserList(userindex).Invent.Object(Slot).Equipped = 0 Then
            Call SendData(ToIndex, userindex, 0, "%J")
            Exit Sub
        End If
        
        If UserList(userindex).flags.Trabajando Then
            Call SendData(ToIndex, userindex, 0, "%K")
            Exit Sub
        End If
        
        Select Case OBJIndex
            Case OBJTYPE_CAÑA, RED_PESCA
                Call SendData(ToIndex, userindex, 0, "T01" & Pesca)
            Case HACHA_LEÑADOR
                Call SendData(ToIndex, userindex, 0, "T01" & Talar)
            Case PIQUETE_MINERO, PICO_EXPERTO
                Call SendData(ToIndex, userindex, 0, "T01" & Mineria)
            Case MARTILLO_HERRERO
                Call SendData(ToIndex, userindex, 0, "T01" & Herreria)
           Case 668 ' hoz LEITO BOTANICA
                Call SendData(ToIndex, userindex, 0, "T01" & Botanica)
            Case SERRUCHO_CARPINTERO
                Call EnviarObjConstruibles(userindex)
                Call SendData(ToIndex, userindex, 0, "SFC")
            Case HILAR_SASTRE
                Call EnviarRopasConstruibles(userindex)
                Call SendData(ToIndex, userindex, 0, "SFS")
        End Select

     Case OBJTYPE_WARP
    
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU" & FONTTYPE_INFO)
            Exit Sub
        End If
        If Not UserList(userindex).flags.TargetNpcTipo = 6 Then
               Call SendData(ToIndex, userindex, 0, "5Y")
               Exit Sub
        Else
               If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 4 Then
                    Call SendData(ToIndex, userindex, 0, "6Y")
                    Exit Sub
               Else
                    If val(Obj.WI) = val(UserList(userindex).POS.Map) Then
                        Call WarpUserChar(userindex, Obj.WMapa, Obj.WX, Obj.WY, True)
                        Call QuitarUserInvItem(userindex, Slot, 1)
                        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_WARP)
                        Call UpdateUserInv(False, userindex, Slot)
                    Else
                        Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "Ese pasaje no te lo he vendido yo, lárgate!" & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
                        Exit Sub
                    End If
               End If
        End If
        
        Case OBJTYPE_PERGAMINOS
            If UserList(userindex).flags.Muerto Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
            End If
            
            If Not ClasePuedeHechizo(userindex, UserList(userindex).Invent.Object(Slot).OBJIndex) Then
                Call SendData(ToIndex, userindex, 0, "||Tu clase no puede aprender este hechizo." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).flags.Hambre = 0 And _
               UserList(userindex).flags.Sed = 0 Then
                Call AgregarHechizo(userindex, Slot)
                Call UpdateUserInv(False, userindex, Slot)
            Else
               Call SendData(ToIndex, userindex, 0, "7F")
            End If
       
       Case OBJTYPE_MINERALES
           If UserList(userindex).flags.Muerto Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
           End If
           Call SendData(ToIndex, userindex, 0, "T01" & FundirMetal)
       
       Case OBJTYPE_INSTRUMENTOS
            If UserList(userindex).flags.Muerto Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
            End If
            Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & Obj.Snd1)
       
       Case OBJTYPE_BARCOS
               If UserList(userindex).flags.Montado = 1 Then Exit Sub

        If ((LegalPos(UserList(userindex).POS.Map, UserList(userindex).POS.x - 1, UserList(userindex).POS.Y, True) Or _
            LegalPos(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1, True) Or _
            LegalPos(UserList(userindex).POS.Map, UserList(userindex).POS.x + 1, UserList(userindex).POS.Y, True) Or _
            LegalPos(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y + 1, True)) And _
            UserList(userindex).flags.Navegando = 0) _
            Or UserList(userindex).flags.Navegando = 1 Then
                Call DoNavega(userindex, CInt(Slot))
        Else
            Call SendData(ToIndex, userindex, 0, "2G")
        End If
           
End Select

End Sub
Sub EnviarArmasConstruibles(userindex As Integer)
Dim i As Integer, cad As String
Dim Descuento As Single

If UserList(userindex).Clase = HERRERO And UserList(userindex).Recompensas(3) = 1 Then
    Descuento = 0.75
Else: Descuento = 1
End If

For i = 1 To UBound(ArmasHerrero)
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    If ObjData(ArmasHerrero(i).Index).SkHerreria <= UserList(userindex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(userindex).Clase) Then
        If ArmasHerrero(i).Recompensa = 0 Or UserList(userindex).Recompensas(2) = 1 Then
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
            cad = cad & ObjData(ArmasHerrero(i).Index).Name & " (" & ObjData(ArmasHerrero(i).Index).MinHIT & "/" & ObjData(ArmasHerrero(i).Index).MaxHIT & ")" & " - (" & Int(val(ObjData(ArmasHerrero(i).Index).LingH * Descuento) * ModMateriales(UserList(userindex).Clase)) & "/" & Int(val(ObjData(ArmasHerrero(i).Index).LingP * Descuento) * ModMateriales(UserList(userindex).Clase)) & "/" & Int(val(ObjData(ArmasHerrero(i).Index).LingO * Descuento) * ModMateriales(UserList(userindex).Clase)) & ")" _
            & "," & ArmasHerrero(i).Index & ","
        End If
    End If
Next

Call SendData(ToIndex, userindex, 0, "LAH" & cad)

End Sub
Sub EnviarObjConstruibles(userindex As Integer)
Dim i As Integer, cad As String, Coste As Integer

For i = 1 To UBound(ObjCarpintero)
'FIXIT: intero(i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    If ObjData(ObjCarpintero(i).Index).SkCarpinteria <= UserList(userindex).Stats.UserSkills(Carpinteria) / ModCarpinteria(UserList(userindex).Clase) Then
        If ObjCarpintero(i).Recompensa = 0 Or (UserList(userindex).Clase = CARPINTERO And UserList(userindex).Recompensas(1) = ObjCarpintero(i).Recompensa) Then
'FIXIT: intero(i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
            Coste = ObjData(ObjCarpintero(i).Index).Madera
'FIXIT: intero(i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
            If UserList(userindex).Clase = CARPINTERO And UserList(userindex).Recompensas(2) = 2 And ObjData(ObjCarpintero(i).Index).ObjType = OBJTYPE_BARCOS Then Coste = Coste * 0.8
'FIXIT: intero(i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: intero(i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: intero(i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
            cad = cad & ObjData(ObjCarpintero(i).Index).Name & " (" & CLng(Coste * ModMadera(UserList(userindex).Clase)) & ") - (" & CLng(val(ObjData(ObjCarpintero(i).Index).MaderaElfica) * ModMadera(UserList(userindex).Clase)) & ")" & "," & ObjCarpintero(i).Index & ","
        End If
    End If
Next

Call SendData(ToIndex, userindex, 0, "OBR" & cad)

End Sub
Sub EnviarRopasConstruibles(userindex As Integer)
Dim PielP As Integer, PielL As Integer, PielO As Integer
Dim N As Integer

Dim i As Integer, cad As String
N = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

For i = 1 To UBound(ObjSastre)
    If ObjData(ObjSastre(i)).SkSastreria <= UserList(userindex).Stats.UserSkills(Sastreria) / ModRopas(UserList(userindex).Clase) Then
        PielP = ObjData(ObjSastre(i)).PielOsoPolar
        PielL = ObjData(ObjSastre(i)).PielLobo
        PielO = ObjData(ObjSastre(i)).PielOsoPardo
        If UserList(userindex).Clase = SASTRE And UserList(userindex).Stats.ELV >= 18 Then
            PielL = PielL * 0.8
            PielO = PielO * 0.8
            PielP = PielP * 0.8
        End If
        cad = cad & ObjData(ObjSastre(i)).Name & " (" & ObjData(ObjSastre(i)).MinDef & "/" & ObjData(ObjSastre(i)).MaxDef & ")" & " - (" & CLng(PielL * ModSastre(UserList(userindex).Clase)) & "/" & CLng(PielO * ModSastre(UserList(userindex).Clase)) & "/" & CLng(PielP * ModSastre(UserList(userindex).Clase)) & ")" & "," & ObjSastre(i) & ","
    End If
Next

Call SendData(ToIndex, userindex, 0, "SAR" & cad)

End Sub


Sub EnviarArmadurasConstruibles(userindex As Integer)
Dim i As Integer, cad As String
Dim Descuento As Single

If UserList(userindex).Clase = HERRERO And UserList(userindex).Recompensas(3) = 1 Then
    Descuento = 0.75
Else: Descuento = 1
End If

For i = 1 To UBound(ArmadurasHerrero)
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    If ObjData(ArmadurasHerrero(i).Index).SkHerreria <= UserList(userindex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(userindex).Clase) Then
        If ArmadurasHerrero(i).Recompensa = 0 Or UserList(userindex).Recompensas(2) = 2 Then
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
            cad = cad & ObjData(ArmadurasHerrero(i).Index).Name & " (" & ObjData(ArmadurasHerrero(i).Index).MinDef & "/" & ObjData(ArmadurasHerrero(i).Index).MaxDef & ")" & " - (" & Int(val(ObjData(ArmadurasHerrero(i).Index).LingH * Descuento) * ModMateriales(UserList(userindex).Clase)) & "/" & Int(val(ObjData(ArmadurasHerrero(i).Index).LingP * Descuento) * ModMateriales(UserList(userindex).Clase)) & "/" & Int(val(ObjData(ArmadurasHerrero(i).Index).LingO * Descuento) * ModMateriales(UserList(userindex).Clase)) & ")" _
            & "," & ArmadurasHerrero(i).Index & ","
        End If
    End If
Next

Call SendData(ToIndex, userindex, 0, "LAR" & cad)


End Sub
Sub EnviarCascosConstruibles(userindex As Integer)
Dim i As Integer, cad As String
Dim Descuento As Single

If UserList(userindex).Clase = HERRERO And UserList(userindex).Recompensas(1) = 2 Then
    Descuento = 0.5
Else: Descuento = 1
End If

For i = 1 To UBound(CascosHerrero)
    If ObjData(CascosHerrero(i)).SkHerreria <= UserList(userindex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(userindex).Clase) Then
        cad = cad & ObjData(CascosHerrero(i)).Name & " (" & ObjData(CascosHerrero(i)).MinDef & "/" & ObjData(CascosHerrero(i)).MaxDef & ")" & " - (" & Int(val(ObjData(CascosHerrero(i)).LingH * Descuento) * ModMateriales(UserList(userindex).Clase)) & "/" & Int(val(ObjData(CascosHerrero(i)).LingP * Descuento) * ModMateriales(UserList(userindex).Clase)) & "/" & Int(val(ObjData(CascosHerrero(i)).LingO * Descuento) * ModMateriales(UserList(userindex).Clase)) & ")" _
        & "," & CascosHerrero(i) & ","
    End If
Next

Call SendData(ToIndex, userindex, 0, "CAS" & cad)

End Sub
Sub EnviarEscudosConstruibles(userindex As Integer)
Dim i As Integer, cad As String
Dim Descuento As Single

If UserList(userindex).Clase = HERRERO And UserList(userindex).Recompensas(1) = 2 Then
    Descuento = 0.5
Else: Descuento = 1
End If

For i = 1 To UBound(EscudosHerrero)
    If ObjData(EscudosHerrero(i)).SkHerreria <= UserList(userindex).Stats.UserSkills(Herreria) / ModHerreriA(UserList(userindex).Clase) Then
        cad = cad & ObjData(EscudosHerrero(i)).Name & " (" & ObjData(EscudosHerrero(i)).MinDef & "/" & ObjData(EscudosHerrero(i)).MaxDef & ") - (" & Int(val(ObjData(EscudosHerrero(i)).LingH * Descuento) * ModMateriales(UserList(userindex).Clase)) & "/" & Int(val(ObjData(EscudosHerrero(i)).LingP * Descuento) * ModMateriales(UserList(userindex).Clase)) & "/" & Int(val(ObjData(EscudosHerrero(i)).LingO * Descuento) * ModMateriales(UserList(userindex).Clase)) & ")" _
        & "," & EscudosHerrero(i) & ","
    End If
Next

Call SendData(ToIndex, userindex, 0, "ESC" & cad)



End Sub
Sub TirarTodo(userindex As Integer)
On Error Resume Next

Call TirarTodosLosItems(userindex)
Call TirarOro(UserList(userindex).Stats.GLD, userindex)

End Sub
Public Function ItemSeCae(Index As Integer) As Boolean

ItemSeCae = (ObjData(Index).Real = 0 And _
            ObjData(Index).Caos = 0 And _
            ObjData(Index).ObjType <> OBJTYPE_LLAVES And _
            ObjData(Index).ObjType <> OBJTYPE_BARCOS And _
            Not ObjData(Index).NoSeCae)

End Function


Sub SaleCarcelPique(userindex As Integer)
Dim i As Byte
Dim Num As Integer
Dim ItemIndex As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(userindex).Invent.Object(i).OBJIndex
    If ItemIndex Then
        If ItemSeCae(ItemIndex) Then
        Num = UserList(userindex).Invent.Object(i).Amount
        If UserList(userindex).Invent.Object(i).Equipped = 1 Then Call Desequipar(userindex, i)
        Call QuitarVariosItem(userindex, i, Num)
        End If
    End If
Next i
Call UpdateUserInv(True, userindex, 1)
End Sub

Sub TirarTodosLosItems(userindex As Integer)

Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer
Dim PosibilidadesZafa As Integer
Dim ZafaMinerales As Boolean

If UserList(userindex).Clase = PIRATA And UserList(userindex).Recompensas(2) = 1 And CInt(RandomNumber(1, 10)) <= 1 Then Exit Sub

If UserList(userindex).Clase = MINERO Then
    If UserList(userindex).Recompensas(1) = 2 Then PosibilidadesZafa = 2
    If UserList(userindex).Recompensas(3) = 2 Then PosibilidadesZafa = PosibilidadesZafa + 3
    ZafaMinerales = CInt(RandomNumber(1, 10)) <= PosibilidadesZafa
End If

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(userindex).Invent.Object(i).OBJIndex
    If ItemIndex Then
        If ItemSeCae(ItemIndex) And Not (ObjData(ItemIndex).ObjType = OBJTYPE_MINERALES And ZafaMinerales) Then
            NuevaPos.x = 0
            NuevaPos.Y = 0
            Call Tilelibre(UserList(userindex).POS, NuevaPos)
            If NuevaPos.x <> 0 And NuevaPos.Y Then
                If MapData(NuevaPos.Map, NuevaPos.x, NuevaPos.Y).OBJInfo.OBJIndex = 0 Then Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.x, NuevaPos.Y)
            End If
        End If
  End If
  
Next

End Sub
Function ItemNewbie(ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function
Sub TirarTodosLosItemsNoNewbies(userindex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
  ItemIndex = UserList(userindex).Invent.Object(i).OBJIndex
  If ItemIndex Then
         If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                NuevaPos.x = 0
                NuevaPos.Y = 0
                Tilelibre UserList(userindex).POS, NuevaPos
                If NuevaPos.x <> 0 And NuevaPos.Y Then
                    If MapData(NuevaPos.Map, NuevaPos.x, NuevaPos.Y).OBJInfo.OBJIndex = 0 Then Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.x, NuevaPos.Y)
                End If
         End If
         
  End If
Next

End Sub
