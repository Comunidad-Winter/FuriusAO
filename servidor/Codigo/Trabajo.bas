Attribute VB_Name = "Trabajo"
Option Explicit

Public Sub DoOcultarse(userindex As Integer)
On Error GoTo errhandler
Dim Suerte As Integer

Suerte = 50 - 0.35 * UserList(userindex).Stats.UserSkills(Ocultarse)

If TiempoTranscurrido(UserList(userindex).Counters.LastOculto) < 0.5 Then Exit Sub
UserList(userindex).Counters.LastOculto = Timer

If UserList(userindex).Clase = CAZADOR Or UserList(userindex).Clase = ASESINO Or UserList(userindex).Clase = LADRON Then Suerte = Suerte - 5

If CInt(RandomNumber(1, Suerte)) <= 5 Then
    UserList(userindex).flags.Oculto = 1
    UserList(userindex).flags.Invisible = 1
    Call SendData(ToMap, 0, UserList(userindex).POS.Map, ("V3" & UserList(userindex).Char.CharIndex & ",1"))
    Call SendData(ToIndex, userindex, 0, "V7")
    Call SubirSkill(userindex, Ocultarse, 15)
Else: Call SendData(ToIndex, userindex, 0, "EN")
End If

Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub
Public Sub DoNavega(userindex As Integer, Slot As Integer)
Dim Barco As ObjData, Skill As Byte

Barco = ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex)

If UserList(userindex).Clase <> PIRATA And UserList(userindex).Clase <> PESCADOR Then
    Skill = Barco.MinSkill * 2
ElseIf UserList(userindex).Invent.Object(Slot).OBJIndex = 474 Then
    Skill = 40
Else: Skill = Barco.MinSkill
End If

If UserList(userindex).Stats.UserSkills(Navegacion) < Skill Then
    If Skill <= 100 Then
        Call SendData(ToIndex, userindex, 0, "!7" & Skill)
    Else: Call SendData(ToIndex, userindex, 0, "||Esta embarcación solo puede ser usada por piratas." & FONTTYPE_INFO)
    End If
    Exit Sub
End If

If UserList(userindex).Stats.ELV < 18 Then
    Call SendData(ToIndex, userindex, 0, "||Debes ser nivel 18 o superior para poder navegar." & FONTTYPE_INFO)
    Exit Sub
End If

UserList(userindex).Invent.BarcoObjIndex = UserList(userindex).Invent.Object(Slot).OBJIndex
UserList(userindex).Invent.BarcoSlot = Slot
           
If UserList(userindex).flags.Navegando = 0 Then
    UserList(userindex).Char.Head = 0
    
    If UserList(userindex).flags.Muerto = 0 Then
        UserList(userindex).Char.Body = Barco.Ropaje
    Else
        UserList(userindex).Char.Body = iFragataFantasmal
    End If
    
    UserList(userindex).Char.ShieldAnim = NingunEscudo
    UserList(userindex).Char.WeaponAnim = NingunArma
    UserList(userindex).Char.CascoAnim = NingunCasco
    UserList(userindex).flags.Navegando = 1
Else
    UserList(userindex).flags.Navegando = 0
    
    If UserList(userindex).flags.Muerto = 0 Then
        UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
        
        If UserList(userindex).Invent.ArmourEqpObjIndex Then
            UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else: Call DarCuerpoDesnudo(userindex)
        End If
            
        If UserList(userindex).Invent.EscudoEqpObjIndex Then _
            UserList(userindex).Char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(userindex).Invent.WeaponEqpObjIndex Then _
            UserList(userindex).Char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(userindex).Invent.CascoEqpObjIndex Then _
            UserList(userindex).Char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(userindex).Char.Body = iCuerpoMuerto
        UserList(userindex).Char.Head = iCabezaMuerto
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserCharB(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
Call SendData(ToIndex, userindex, 0, "NAVEG")

End Sub
Public Sub FundirMineral(userindex As Integer)

If UserList(userindex).flags.TargetObjInvIndex Then
    If ObjData(UserList(userindex).flags.TargetObjInvIndex).MinSkill <= UserList(userindex).Stats.UserSkills(Mineria) / ModFundicion(UserList(userindex).Clase) Then
         Call DoLingotes(userindex)
    Else: Call SendData(ToIndex, userindex, 0, "!8")
    End If
End If

End Sub
Function TieneObjetos(ItemIndex As Integer, Cant As Integer, userindex As Integer) As Boolean
Dim i As Byte
Dim Total As Long

For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(i).OBJIndex = ItemIndex Then
        Total = Total + UserList(userindex).Invent.Object(i).Amount
    End If
Next

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function
Function QuitarObjetos(ItemIndex As Integer, Cant As Integer, userindex As Integer) As Boolean
Dim i As Byte

For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(i).OBJIndex = ItemIndex Then
        
        Call Desequipar(userindex, i)
        
        UserList(userindex).Invent.Object(i).Amount = UserList(userindex).Invent.Object(i).Amount - Cant
        If (UserList(userindex).Invent.Object(i).Amount <= 0) Then
            Cant = Abs(UserList(userindex).Invent.Object(i).Amount)
            UserList(userindex).Invent.Object(i).Amount = 0
            UserList(userindex).Invent.Object(i).OBJIndex = 0
        Else
            Cant = 0
        End If
        
        Call UpdateUserInv(False, userindex, i)
        
        If (Cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next

End Function
Sub HerreroQuitarMateriales(userindex As Integer, ItemIndex As Integer, cantT As Integer)
Dim Descuento As Single

Descuento = 1

If UserList(userindex).Clase = HERRERO Then
    If UserList(userindex).Recompensas(1) = 1 And ObjData(ItemIndex).SubTipo <> OBJTYPE_CASCO And ObjData(ItemIndex).SubTipo <> OBJTYPE_ESCUDO Then
        If CInt(RandomNumber(1, 4)) <= 1 Then Descuento = 0.5
    ElseIf UserList(userindex).Recompensas(1) = 2 And (ObjData(ItemIndex).SubTipo = OBJTYPE_CASCO Or ObjData(ItemIndex).SubTipo = OBJTYPE_ESCUDO) Then
        Descuento = 0.5
    End If
    Descuento = Descuento * (1 - 0.25 * Buleano(UserList(userindex).Recompensas(3) = 1 And ObjData(ItemIndex).SubTipo <> OBJTYPE_CASCO And ObjData(ItemIndex).SubTipo <> OBJTYPE_ESCUDO))
End If

If ObjData(ItemIndex).LingH Then Call QuitarObjetos(LingoteHierro, Descuento * Int(ObjData(ItemIndex).LingH * ModMateriales(UserList(userindex).Clase) * cantT), userindex)
If ObjData(ItemIndex).LingP Then Call QuitarObjetos(LingotePlata, Descuento * Int(ObjData(ItemIndex).LingP * ModMateriales(UserList(userindex).Clase) * cantT), userindex)
If ObjData(ItemIndex).LingO Then Call QuitarObjetos(LingoteOro, Descuento * Int(ObjData(ItemIndex).LingO * ModMateriales(UserList(userindex).Clase) * cantT), userindex)

End Sub
Sub CarpinteroQuitarMateriales(userindex As Integer, ItemIndex As Integer, cantT As Integer)
Dim Descuento As Single

cantT = Maximo(1, cantT)

If UserList(userindex).Clase = CARPINTERO And UserList(userindex).Recompensas(2) = 2 And ObjData(ItemIndex).ObjType = OBJTYPE_BARCOS Then
    Descuento = 0.8
Else
    Descuento = 1
End If

If ObjData(ItemIndex).Madera Then
    Call QuitarObjetos(Leña, CInt(Descuento * ObjData(ItemIndex).Madera * ModMadera(UserList(userindex).Clase) * cantT), userindex)
End If

If ObjData(ItemIndex).MaderaElfica Then
    Call QuitarObjetos(LeñaElfica, CInt(Descuento * ObjData(ItemIndex).MaderaElfica * ModMadera(UserList(userindex).Clase) * cantT), userindex)
End If

End Sub
Sub AlquimiaQuitarMateriales(ByVal userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)
    If ObjData(ItemIndex).Hierbas > 0 Then Call QuitarObjetos(667, ObjData(ItemIndex).Hierbas * Cant, userindex)
End Sub
 Function AlquimiaTieneMateriales(ByVal userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer) As Boolean
    
    If ObjData(ItemIndex).Hierbas > 0 Then
            If Not TieneObjetos(667, ObjData(ItemIndex).Hierbas * Cant, userindex) Then
                    Call SendData(ToIndex, userindex, 0, "||No tenes suficientes hierbas." & FONTTYPE_INFO)
                    AlquimiaTieneMateriales = False
                    Exit Function
            End If
    End If
    
    AlquimiaTieneMateriales = True

End Function
Function CarpinteroTieneMateriales(userindex As Integer, ItemIndex As Integer, cantT As Integer) As Boolean
Dim Descuento As Single

cantT = Maximo(1, cantT)

If UserList(userindex).Clase = CARPINTERO And UserList(userindex).Recompensas(2) = 2 And ObjData(ItemIndex).ObjType = OBJTYPE_BARCOS Then
    Descuento = 0.8
Else
    Descuento = 1
End If

If ObjData(ItemIndex).Madera Then
    If Not TieneObjetos(Leña, CInt(Descuento * ObjData(ItemIndex).Madera * ModMadera(UserList(userindex).Clase) * cantT), userindex) Then
        Call SendData(ToIndex, userindex, 0, "!9")
        CarpinteroTieneMateriales = False
        Exit Function
    End If
End If
    
If ObjData(ItemIndex).MaderaElfica Then
    If Not TieneObjetos(LeñaElfica, CInt(Descuento * ObjData(ItemIndex).MaderaElfica * ModMadera(UserList(userindex).Clase) * cantT), userindex) Then
        Call SendData(ToIndex, userindex, 0, "!9")
        CarpinteroTieneMateriales = False
        Exit Function
    End If
End If
    
CarpinteroTieneMateriales = True

End Function
Function Piel(userindex As Integer, Tipo As Byte, Obj As Integer) As Integer

Select Case Tipo
    Case 1
        Piel = ObjData(Obj).PielLobo
        If UserList(userindex).Clase = SASTRE And UserList(userindex).Stats.ELV >= 18 Then Piel = Piel * 0.8
    Case 2
        Piel = ObjData(Obj).PielOsoPardo
        If UserList(userindex).Clase = SASTRE And UserList(userindex).Stats.ELV >= 18 Then Piel = Piel * 0.8
    Case 3
        Piel = ObjData(Obj).PielOsoPolar
        If UserList(userindex).Clase = SASTRE And UserList(userindex).Stats.ELV >= 18 Then Piel = Piel * 0.8
End Select

End Function
Function SastreTieneMateriales(userindex As Integer, ItemIndex As Integer, cantT As Integer) As Boolean
Dim PielL As Integer, PielO As Integer, PielP As Integer
cantT = Maximo(1, cantT)

PielL = ObjData(ItemIndex).PielLobo
PielO = ObjData(ItemIndex).PielOsoPardo
PielP = ObjData(ItemIndex).PielOsoPolar

If UserList(userindex).Clase = SASTRE And UserList(userindex).Stats.ELV >= 18 Then
    PielL = 0.8 * PielL
    PielO = 0.8 * PielO
    PielP = 0.8 * PielP
End If

If PielL Then
    If Not TieneObjetos(PLobo, CInt(PielL * ModSastre(UserList(userindex).Clase)) * cantT, userindex) Then
        Call SendData(ToIndex, userindex, 0, "0A")
        SastreTieneMateriales = False
        Exit Function
    End If
End If

If PielO Then
    If Not TieneObjetos(POsoPardo, CInt(PielO * ModSastre(UserList(userindex).Clase)) * cantT, userindex) Then
        Call SendData(ToIndex, userindex, 0, "0A")
        SastreTieneMateriales = False
        Exit Function
    End If
End If
    
If PielP Then
    If Not TieneObjetos(POsoPolar, CInt(PielP * ModSastre(UserList(userindex).Clase)) * cantT, userindex) Then
        Call SendData(ToIndex, userindex, 0, "0A")
        SastreTieneMateriales = False
        Exit Function
    End If
End If
    
SastreTieneMateriales = True

End Function
Sub SastreQuitarMateriales(userindex As Integer, ItemIndex As Integer, cantT As Integer)
Dim PielL As Integer, PielO As Integer, PielP As Integer

PielL = ObjData(ItemIndex).PielLobo
PielO = ObjData(ItemIndex).PielOsoPardo
PielP = ObjData(ItemIndex).PielOsoPolar

If UserList(userindex).Clase = SASTRE And UserList(userindex).Stats.ELV >= 18 Then
    PielL = 0.8 * PielL
    PielO = 0.8 * PielO
    PielP = 0.8 * PielP
End If

If PielL Then Call QuitarObjetos(PLobo, CInt(PielL * ModSastre(UserList(userindex).Clase)) * cantT, userindex)
If PielO Then Call QuitarObjetos(POsoPardo, CInt(PielO * ModSastre(UserList(userindex).Clase)) * cantT, userindex)
If PielP Then Call QuitarObjetos(POsoPolar, CInt(PielP * ModSastre(UserList(userindex).Clase)) * cantT, userindex)

End Sub
Public Sub SastreConstruirItem(userindex As Integer, ItemIndex As Integer, cantT As Integer)

If SastreTieneMateriales(userindex, ItemIndex, cantT) And _
   UserList(userindex).Stats.UserSkills(Sastreria) / ModRopas(UserList(userindex).Clase) >= _
   ObjData(ItemIndex).SkSastreria And _
   PuedeConstruirSastre(ItemIndex, userindex) And _
   UserList(userindex).Invent.HerramientaEqpObjIndex = HILAR_SASTRE Then
        
    Call SastreQuitarMateriales(userindex, ItemIndex, cantT)
    Call SendData(ToIndex, userindex, 0, "0C")
    
    Dim MiObj As Obj
    MiObj.Amount = Maximo(1, cantT)
    MiObj.OBJIndex = ItemIndex
    
    If Not MeterItemEnInventario(userindex, MiObj) Then Call TirarItemAlPiso(UserList(userindex).POS, MiObj)
    
    Call CheckUserLevel(userindex)

    Call SubirSkill(userindex, Sastreria, 5)

Else
    Call SendData(ToIndex, userindex, 0, "0D")

End If

End Sub

Public Function PuedeConstruirSastre(ItemIndex As Integer, userindex As Integer) As Boolean
Dim i As Long
Dim N As Integer

N = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

For i = 1 To UBound(ObjSastre)
    If ObjSastre(i) = ItemIndex Then
        PuedeConstruirSastre = True
        Exit Function
    End If
Next

PuedeConstruirSastre = False

End Function

Function HerreroTieneMateriales(userindex As Integer, ItemIndex As Integer, cantT As Integer) As Boolean
Dim Descuento As Single

Descuento = 1

If UserList(userindex).Clase = HERRERO Then
    If UserList(userindex).Recompensas(1) = 1 And ObjData(ItemIndex).SubTipo <> OBJTYPE_CASCO And ObjData(ItemIndex).SubTipo <> OBJTYPE_ESCUDO Then
        If CInt(RandomNumber(1, 4)) <= 1 Then Descuento = 0.5
    ElseIf UserList(userindex).Recompensas(1) = 2 And (ObjData(ItemIndex).SubTipo = OBJTYPE_CASCO Or ObjData(ItemIndex).SubTipo = OBJTYPE_ESCUDO) Then
        Descuento = 0.5
    End If
    Descuento = Descuento * (1 - 0.25 * Buleano(UserList(userindex).Recompensas(3) = 1 And ObjData(ItemIndex).SubTipo <> OBJTYPE_CASCO And ObjData(ItemIndex).SubTipo <> OBJTYPE_ESCUDO))
End If

If ObjData(ItemIndex).LingH Then
    If Not TieneObjetos(LingoteHierro, Descuento * Int(ObjData(ItemIndex).LingH * ModMateriales(UserList(userindex).Clase) * cantT), userindex) Then
        Call SendData(ToIndex, userindex, 0, "0E")
        HerreroTieneMateriales = False
        Exit Function
    End If
End If
If ObjData(ItemIndex).LingP Then
    If Not TieneObjetos(LingotePlata, Descuento * Int(ObjData(ItemIndex).LingP * ModMateriales(UserList(userindex).Clase) * cantT), userindex) Then
        Call SendData(ToIndex, userindex, 0, "0F")
        HerreroTieneMateriales = False
        Exit Function
    End If
End If
If ObjData(ItemIndex).LingO Then
    If Not TieneObjetos(LingoteOro, Descuento * Int(ObjData(ItemIndex).LingO * ModMateriales(UserList(userindex).Clase) * cantT), userindex) Then
        Call SendData(ToIndex, userindex, 0, "0G")
        HerreroTieneMateriales = False
        Exit Function
    End If
End If
HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(userindex As Integer, ItemIndex As Integer, cantT As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(userindex, ItemIndex, cantT) And UserList(userindex).Stats.UserSkills(Herreria) >= ObjData(ItemIndex).SkHerreria * ModHerreriA(UserList(userindex).Clase)
End Function
Public Function PuedeConstruirHerreria(ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    If ArmasHerrero(i).Index = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next

For i = 1 To UBound(ArmadurasHerrero)
'FIXIT: i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    If ArmadurasHerrero(i).Index = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next

For i = 1 To UBound(CascosHerrero)
    If CascosHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next

For i = 1 To UBound(EscudosHerrero)
    If EscudosHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next

PuedeConstruirHerreria = False

End Function
Public Sub HerreroConstruirItem(userindex As Integer, ItemIndex As Integer, cantT As Integer)

If cantT > 10 Then
    Call SendData(ToIndex, userindex, 0, "0H")
    Exit Sub
End If

If PuedeConstruir(userindex, ItemIndex, cantT) And PuedeConstruirHerreria(ItemIndex) Then
    Call HerreroQuitarMateriales(userindex, ItemIndex, cantT)
    
    Select Case ObjData(ItemIndex).ObjType
        Case OBJTYPE_WEAPON
            Call SendData(ToIndex, userindex, 0, "0I")
        Case OBJTYPE_ESCUDO
            Call SendData(ToIndex, userindex, 0, "0L")
        Case OBJTYPE_CASCO
            Call SendData(ToIndex, userindex, 0, "0K")
        Case OBJTYPE_ARMOUR
            Call SendData(ToIndex, userindex, 0, "0J")
    End Select
    cantT = cantT * (1 + Buleano(CInt(RandomNumber(1, 10)) <= 1 And UserList(userindex).Clase = HERRERO And UserList(userindex).Recompensas(3) = 2))
    Dim MiObj As Obj
    MiObj.Amount = cantT
    MiObj.OBJIndex = ItemIndex
    If Not MeterItemEnInventario(userindex, MiObj) Then Call TirarItemAlPiso(UserList(userindex).POS, MiObj)

    Call CheckUserLevel(userindex)
    Call SubirSkill(userindex, Herreria, 5)
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & MARTILLOHERRERO)
    Else

End If

End Sub
Public Function PuedeConstruirCarpintero(ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
'FIXIT: intero(i).Index property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
    If ObjCarpintero(i).Index = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next
PuedeConstruirCarpintero = False

End Function
Public Sub CarpinteroConstruirItem(userindex As Integer, ItemIndex As Integer, cantT As Integer)

If CarpinteroTieneMateriales(userindex, ItemIndex, cantT) And _
   UserList(userindex).Stats.UserSkills(Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(userindex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then

    Call CarpinteroQuitarMateriales(userindex, ItemIndex, cantT)
    Call SendData(ToIndex, userindex, 0, "0M")
    
    Dim MiObj As Obj
    If UserList(userindex).Clase = CARPINTERO And UserList(userindex).Recompensas(2) = 1 And ObjData(ItemIndex).ObjType = OBJTYPE_FLECHAS Then cantT = cantT * 2
    MiObj.Amount = cantT
    MiObj.OBJIndex = ItemIndex

    If Not MeterItemEnInventario(userindex, MiObj) Then Call TirarItemAlPiso(UserList(userindex).POS, MiObj)

    
    Call CheckUserLevel(userindex)

    Call SubirSkill(userindex, Carpinteria, 5)
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & LABUROCARPINTERO)
End If

End Sub
Public Sub AlquimiaConstruirItem(ByVal userindex As Integer, ByVal ItemIndex As Integer, ByVal Cant As Integer)

If AlquimiaTieneMateriales(userindex, ItemIndex, Cant) And _
   UserList(userindex).Stats.UserSkills(Alquimia) >= _
   ObjData(ItemIndex).SkAlquimia And _
   PuedeConstruirBotanica(ItemIndex) And _
   UserList(userindex).Invent.HerramientaEqpObjIndex = 669 Then ' cacerola

    Call AlquimiaQuitarMateriales(userindex, ItemIndex, Cant)
    Call SendData(ToIndex, userindex, 0, "||Has mezclado las hierbas!" & FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = Cant
    MiObj.OBJIndex = ItemIndex
    If Not MeterItemEnInventario(userindex, MiObj) Then
                    Call TirarItemAlPiso(UserList(userindex).POS, MiObj)
    End If
    
    Call SubirSkill(userindex, Alquimia)
    Call UpdateUserInv(True, userindex, 0)
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & LABUROCARPINTERO)
End If

End Sub
Public Function PuedeConstruirBotanica(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjBotanica)
    If ObjBotanica(i) = ItemIndex Then
        PuedeConstruirBotanica = True
        Exit Function
    End If
Next i
PuedeConstruirBotanica = False

End Function
Public Sub DoLingotes(userindex As Integer)
Dim Minimo As Integer

Select Case ObjData(UserList(userindex).flags.TargetObjInvIndex).LingoteIndex
    Case LingoteHierro
        Minimo = 6
    Case LingotePlata
        Minimo = 18
    Case LingoteOro
        Minimo = 34
End Select

If UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvslot).Amount < Minimo Then
    Call SendData(ToIndex, userindex, 0, "M3")
    Exit Sub
End If

Dim nPos As WorldPos
Dim MiObj As Obj

MiObj.Amount = UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvslot).Amount / Minimo
MiObj.OBJIndex = ObjData(UserList(userindex).flags.TargetObjInvIndex).LingoteIndex

If Not MeterItemEnInventario(userindex, MiObj) Then Call TirarItemAlPiso(UserList(userindex).POS, MiObj)

UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvslot).Amount = 0
UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvslot).OBJIndex = 0

Call UpdateUserInv(False, userindex, UserList(userindex).flags.TargetObjInvslot)
Call SendData(ToIndex, userindex, 0, "M1")

End Sub
Function ModFundicion(Clase As Byte) As Single

Select Case (Clase)
    Case MINERO, HERRERO
        ModFundicion = 1
    Case TRABAJADOR, EXPERTO_MINERALES
        ModFundicion = 2.5
    Case Else
        ModFundicion = 3
End Select

End Function
Function ModHerreriA(Clase As Byte) As Single

Select Case (Clase)
    Case HERRERO
        ModHerreriA = 1
    Case Else
        ModHerreriA = 3
End Select

End Function

Function ModCarpinteria(Clase As Byte) As Single

Select Case (Clase)
    Case CARPINTERO
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function
Function ModMateriales(Clase As Byte) As Single

Select Case (Clase)
    Case HERRERO
        ModMateriales = 1
    Case Else
        ModMateriales = 3
End Select

End Function
Function ModMadera(Clase As Byte) As Double

Select Case (Clase)
    Case CARPINTERO
        ModMadera = 1
    Case Else
        ModMadera = 3
End Select

End Function
Function ModSastre(Clase As Byte) As Double

Select Case (Clase)
    Case SASTRE
        ModSastre = 1
    Case Else
        ModSastre = 3
End Select

End Function
Function ModRopas(Clase As Byte) As Double

Select Case (Clase)
    Case SASTRE
        ModRopas = 1
    Case Else
        ModRopas = 3
End Select

End Function
Function FreeMascotaIndex(userindex As Integer) As Integer
Dim j As Integer

For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(j) = 0 Then
        FreeMascotaIndex = j
        Exit Function
    End If
Next

End Function
Sub DoDomar(userindex As Integer, NpcIndex As Integer)


If UserList(userindex).NroMascotas < 3 Then
    If Npclist(NpcIndex).MaestroUser = userindex Then
        Call SendData(ToIndex, userindex, 0, "0N")
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(ToIndex, userindex, 0, "0Ñ")
        Exit Sub
    End If
    
    If Npclist(NpcIndex).flags.Domable <= UserList(userindex).Stats.UserSkills(Domar) Then
        Dim Index As Integer
        'If TransformarMascota(NpcIndex, userindex) Then Exit Sub
        UserList(userindex).NroMascotas = UserList(userindex).NroMascotas + 1
        Index = FreeMascotaIndex(userindex)
        UserList(userindex).MascotasIndex(Index) = NpcIndex
        UserList(userindex).MascotasType(Index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = userindex
        
        Call QuitarNPCDeLista(Npclist(NpcIndex).Numero, Npclist(NpcIndex).POS.Map)
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(ToIndex, userindex, 0, "0O")
        Call SubirSkill(userindex, Domar)
        
    Else
    
        If UserList(userindex).Clase = Druida And UserList(userindex).Recompensas(3) = 2 Then
            If UserList(userindex).NroMascotas < 2 Then
                'If TransformarMascota(NpcIndex, userindex) Then Exit Sub
                Dim i As Integer
                For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
                    If UserList(userindex).MascotasIndex(i) Then
                        If Npclist(NpcIndex).Name = Npclist(UserList(userindex).MascotasIndex(i)).Name Then
                            Call SendData(ToIndex, userindex, 0, "||No puedes domar dos criaturas iguales." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                    End If
                Next
        
        
                UserList(userindex).NroMascotas = UserList(userindex).NroMascotas + 1
                Index = FreeMascotaIndex(userindex)
                UserList(userindex).MascotasIndex(Index) = NpcIndex
                UserList(userindex).MascotasType(Index) = Npclist(NpcIndex).Numero
                Npclist(NpcIndex).MaestroUser = userindex
                Call QuitarNPCDeLista(Npclist(NpcIndex).Numero, Npclist(NpcIndex).POS.Map)
                Call FollowAmo(NpcIndex)
                Call SendData(ToIndex, userindex, 0, "0O")
                Call SubirSkill(userindex, Domar)
                Exit Sub
            End If
        End If
    
        Call SendData(ToIndex, userindex, 0, "||Necesitas " & Npclist(NpcIndex).flags.Domable & " puntos para domar a esta criatura. " & FONTTYPE_INFO)
        
    End If
Else
    Call SendData(ToIndex, userindex, 0, "0Q")
End If

End Sub
Sub DoAdminInvisible(userindex As Integer)

If UserList(userindex).flags.AdminInvisible = 0 Then
    UserList(userindex).flags.AdminInvisible = 1
    UserList(userindex).flags.Invisible = 1
    Call SendData(ToMap, 0, UserList(userindex).POS.Map, ("V3" & UserList(userindex).Char.CharIndex & ",1"))
    Call SendData(ToMap, 0, UserList(userindex).POS.Map, "QDL" & UserList(userindex).Char.CharIndex)
Else
    UserList(userindex).flags.AdminInvisible = 0
    UserList(userindex).flags.Invisible = 0
    Call SendData(ToMap, 0, UserList(userindex).POS.Map, ("V3" & UserList(userindex).Char.CharIndex & ",0"))
End If
    
End Sub
Sub TratarDeHacerFogata(Map As Integer, x As Integer, Y As Integer, userindex As Integer)
Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj, nPos As WorldPos

If Not LegalPos(Map, x, Y) Then Exit Sub
nPos.Map = Map
nPos.x = x
nPos.Y = Y

If Not MapInfo(UserList(userindex).POS.Map).Pk Then
    Call SendData(ToIndex, userindex, 0, "||No puedes hacer fogatas en zonas seguras." & FONTTYPE_WARNING)
    Exit Sub
End If

If Distancia(nPos, UserList(userindex).POS) > 4 Then
    Call SendData(ToIndex, userindex, 0, "DL")
    Exit Sub
End If

If MapData(Map, x, Y).OBJInfo.Amount < 3 Then
    Call SendData(ToIndex, userindex, 0, "0R")
    Exit Sub
End If

If UserList(userindex).Stats.UserSkills(Supervivencia) > 1 And UserList(userindex).Stats.UserSkills(Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 6 And UserList(userindex).Stats.UserSkills(Supervivencia) <= 10 Then
    Suerte = 2
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 10 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.OBJIndex = FOGATA_APAG
    Obj.Amount = MapData(Map, x, Y).OBJInfo.Amount / 3
    
    If Obj.Amount > 1 Then
        Call SendData(ToIndex, userindex, 0, "0S" & Obj.Amount)
    Else
        Call SendData(ToIndex, userindex, 0, "0T")
    End If
    
    Call MakeObj(ToMap, 0, Map, Obj, Map, x, Y)
    
    Dim Fogatita As New cGarbage
    Fogatita.Map = Map
    Fogatita.x = x
    Fogatita.Y = Y
    Call TrashCollector.Add(Fogatita)
    
Else
    Call SendData(ToIndex, userindex, 0, "0U")
End If

Call SubirSkill(userindex, Supervivencia)


End Sub
Public Sub DoTalar(userindex As Integer, Elfico As Boolean)
On Error GoTo errhandler
Dim MiObj As Obj
Dim Factor As Integer
Dim Esfuerzo As Integer

If UserList(userindex).Clase = TALADOR Then
    Esfuerzo = EsfuerzoTalarLeñador
Else
    Esfuerzo = EsfuerzoTalarGeneral
End If

If UserList(userindex).Stats.MinSta >= Esfuerzo Then
    Call QuitarSta(userindex, Esfuerzo)
    Call SendUserSTA(userindex)
Else
    Call SendData(ToIndex, userindex, 0, "9E")
    Exit Sub
End If

If Elfico Then
    MiObj.OBJIndex = LeñaElfica
    Factor = 6
Else
    MiObj.OBJIndex = Leña
    Factor = 5
End If



If UserList(userindex).Clase = TALADOR Then
    MiObj.Amount = Fix(4 + ((0.29 + 0.07 * Buleano(UserList(userindex).Recompensas(1) = 1)) * UserList(userindex).Stats.UserSkills(Talar)))
Else: MiObj.Amount = 1
End If

MiObj.Amount = MiObj.Amount * 5

If Not MeterItemEnInventario(userindex, MiObj) Then Call TirarItemAlPiso(UserList(userindex).POS, MiObj)

Call SendData(ToPCArea, CInt(userindex), UserList(userindex).POS.Map, "TW" & SOUND_TALAR)
Call SubirSkill(userindex, Talar, 5)

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub
'BOTANICA LEITO
Public Sub DoRecolectar(ByVal userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim Res As Integer


If UserList(userindex).Clase = "Druida" Then
    Call QuitarSta(userindex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(userindex, EsfuerzoTalarGeneral)
End If

If UserList(userindex).Stats.UserSkills(Botanica) <= 10 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 20 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 30 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 40 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 50 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 60 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 70 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 80 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 90 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 81 Then
                    Suerte = 13
ElseIf UserList(userindex).Stats.UserSkills(Botanica) <= 100 _
   And UserList(userindex).Stats.UserSkills(Botanica) >= 91 Then
                    Suerte = 7
End If
Res = RandomNumber(1, Suerte)

If Res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(userindex).Clase = "Druida" Then
        MiObj.Amount = RandomNumber(1, 5)
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.OBJIndex = 667 'objeto hierba
    
    
    If Not MeterItemEnInventario(userindex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(userindex).POS, MiObj)
        
    End If
    
    Call SendData(ToIndex, userindex, 0, "||¡Has conseguido algunas hierbas!" & FONTTYPE_INFO)
    
Else
    Call SendData(ToIndex, userindex, 0, "||¡No has obtenido hierbas!" & FONTTYPE_INFO)
End If

Call SubirSkill(userindex, Botanica)

Exit Sub

errhandler:
    Call LogError("Error en DoRecolectar")

End Sub 'BOTANICA LEITO

Public Sub DoPescar(userindex As Integer)
On Error GoTo errhandler
Dim Esfuerzo As Integer
Dim MiObj As Obj

If UserList(userindex).Clase = PESCADOR Then
    Esfuerzo = EsfuerzoPescarPescador
Else
    Esfuerzo = EsfuerzoPescarGeneral
End If

If UserList(userindex).Stats.MinSta >= Esfuerzo Then
    Call QuitarSta(userindex, Esfuerzo)
    Call SendUserSTA(userindex)
Else
    Call SendData(ToIndex, userindex, 0, "9E")
    Exit Sub
End If

MiObj.OBJIndex = Pescado


If UserList(userindex).Clase = PESCADOR Then
    If UserList(userindex).Recompensas(1) = 2 And UserList(userindex).flags.Navegando = 1 And UserList(userindex).Invent.HerramientaEqpObjIndex = RED_PESCA And CInt(RandomNumber(1, 10)) <= 1 Then MiObj.OBJIndex = PescadoCaro + CInt(RandomNumber(1, 3))
    MiObj.Amount = Fix(4 + ((0.29 + 0.07 * Buleano(UserList(userindex).Recompensas(1) = 1)) * UserList(userindex).Stats.UserSkills(Pesca)))
    If UserList(userindex).Invent.HerramientaEqpObjIndex <> RED_PESCA Then MiObj.Amount = MiObj.Amount / 2
Else: MiObj.Amount = 1
End If

MiObj.Amount = MiObj.Amount + 5

Call SubirSkill(userindex, Pesca, 5)
If Not MeterItemEnInventario(userindex, MiObj) Then Call TirarItemAlPiso(UserList(userindex).POS, MiObj)
Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SOUND_PESCAR)

Exit Sub

errhandler:
    Call LogError("Error en DoPescar")

End Sub
Public Function Buleano(A As Boolean) As Byte

Buleano = -A

End Function
Public Sub DoRobar(LadronIndex As Integer, VictimaIndex As Integer)
Dim Res As Integer
Dim N As Long

If Not PuedeRobar(LadronIndex, VictimaIndex) Then Exit Sub

UserList(LadronIndex).Counters.LastRobo = Timer

Res = RandomNumber(1, 100)

If Res > UserList(LadronIndex).Stats.UserSkills(Robar) \ 10 + 25 * Buleano(UserList(LadronIndex).Clase = LADRON) + 5 * Buleano(UserList(LadronIndex).Clase = LADRON And UserList(LadronIndex).Recompensas(1) = 2) Then
    Call SendData(ToIndex, LadronIndex, 0, "X0")
    Call SendData(ToIndex, VictimaIndex, 0, "Y0" & UserList(LadronIndex).Name)
ElseIf UserList(LadronIndex).Clase = LADRON And TieneObjetosRobables(VictimaIndex) And Res <= 10 * Buleano(UserList(LadronIndex).Recompensas(2) = 2) + 10 * Buleano(UserList(LadronIndex).Recompensas(3) = 2) Then
    Call RobarObjeto(LadronIndex, VictimaIndex)
ElseIf UserList(VictimaIndex).Stats.GLD = 0 Then
    Call SendData(ToIndex, LadronIndex, 0, "W0")
    Call SendData(ToIndex, VictimaIndex, 0, "Y0" & UserList(LadronIndex).Name)
Else
    N = Minimo((1 + 0.1 * Buleano(UserList(LadronIndex).Recompensas(1) = 1 And UserList(LadronIndex).Clase = LADRON)) * (RandomNumber(1, (UserList(LadronIndex).Stats.UserSkills(Robar) * (UserList(VictimaIndex).Stats.ELV / 10) * UserList(LadronIndex).Stats.ELV)) / (10 + 10 * Buleano(Not UserList(LadronIndex).Clase = LADRON))), UserList(VictimaIndex).Stats.GLD)
    
    UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
    Call AddtoVar(UserList(LadronIndex).Stats.GLD, N, MAXORO)
   
    Call SendData(ToIndex, LadronIndex, 0, "U0" & UserList(VictimaIndex).Name & "," & N)
    Call SendData(ToIndex, VictimaIndex, 0, "V0" & UserList(LadronIndex).Name & "," & N)
    
    Call SendUserORO(LadronIndex)
    Call SendUserORO(VictimaIndex)
End If

Call SubirSkill(LadronIndex, Robar)

End Sub
Public Function ObjEsRobable(VictimaIndex As Integer, Slot As Byte) As Boolean
Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).OBJIndex
If OI = 0 Then Exit Function

ObjEsRobable = ObjData(OI).ObjType <> OBJTYPE_LLAVES And _
                ObjData(OI).ObjType <> OBJTYPE_BARCOS And _
                ObjData(OI).Real = 0 And _
                ObjData(OI).Caos = 0 And _
                ObjData(OI).NoSeCae = False

End Function
Public Sub RobarObjeto(LadronIndex As Integer, VictimaIndex As Integer)
Dim IndexRobo As Byte
Dim MiObj As Obj
Dim Num As Byte

Do
    IndexRobo = RandomNumber(1, MAX_INVENTORY_SLOTS)
    If ObjEsRobable(VictimaIndex, IndexRobo) Then Exit Do
Loop

MiObj.OBJIndex = UserList(VictimaIndex).Invent.Object(IndexRobo).OBJIndex

Num = Minimo(RandomNumber(1, 4 + 96 * Buleano(ObjData(MiObj.OBJIndex).ObjType = OBJTYPE_POCIONES)), UserList(VictimaIndex).Invent.Object(IndexRobo).Amount)

If UserList(VictimaIndex).Invent.Object(IndexRobo).Equipped = 1 Then Call Desequipar(VictimaIndex, IndexRobo)

MiObj.Amount = Num

UserList(VictimaIndex).Invent.Object(IndexRobo).Amount = UserList(VictimaIndex).Invent.Object(IndexRobo).Amount - Num
If UserList(VictimaIndex).Invent.Object(IndexRobo).Amount <= 0 Then Call QuitarUserInvItem(VictimaIndex, CByte(IndexRobo), 1)

If Not MeterItemEnInventario(LadronIndex, MiObj) Then Call TirarItemAlPiso(UserList(LadronIndex).POS, MiObj)

Call SendData(ToIndex, LadronIndex, 0, "||Has robado " & ObjData(MiObj.OBJIndex).Name & " (" & MiObj.Amount & ")." & FONTTYPE_INFO)
Call UpdateUserInv(False, VictimaIndex, CByte(IndexRobo))

End Sub
Public Sub DoApuñalar(userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal Daño As Integer)
Dim Prob As Integer

Prob = 20 - 1.2 * UserList(userindex).Stats.UserSkills(Apuñalar) \ 10

'Select Case UserList(UserIndex).Clase
'    Case ASESINO
'        Prob = Prob - 3 - Buleano(UserList(UserIndex).Recompensas(3) = 2)
'    Case BARDO
'        Prob = Prob - 2 - Buleano(UserList(UserIndex).Recompensas(3) = 1)
'End Select
If UserList(userindex).Clase = ASESINO Then Prob = Prob - 3 - Buleano(UserList(userindex).Recompensas(3) = 2)


If RandomNumber(1, Prob) <= 1 Then
    If VictimUserIndex Then
        If UserList(userindex).Clase = ASESINO And UserList(userindex).Recompensas(3) = 1 Then
            Daño = Daño * 1.7
        Else: Daño = Daño * 1.5
        End If
        If Not UserList(VictimUserIndex).flags.Quest And UserList(VictimUserIndex).flags.Privilegios = 0 Then
            UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Daño
            Call SendUserHP(VictimUserIndex)
        End If
        Call SendData(ToIndex, userindex, 0, "5K" & UserList(VictimUserIndex).Name & "," & Daño)
        Call SendData(ToIndex, VictimUserIndex, 0, "5L" & UserList(userindex).Name & "," & Daño)
    ElseIf VictimNpcIndex Then
        Select Case UserList(userindex).Clase
            Case ASESINO
                Daño = Daño * 2
            Case Else
                Daño = Daño * 1.5
        End Select
        Call SendData(ToIndex, userindex, 0, "5M" & Daño)
        Call ExperienciaPorGolpe(userindex, VictimNpcIndex, Daño)
        Call VerNPCMuere(VictimNpcIndex, Daño, userindex)
    End If
Else
    Call SendData(ToIndex, userindex, 0, "5N")
End If

End Sub
Public Sub QuitarSta(userindex As Integer, Cantidad As Integer)

If UserList(userindex).flags.Quest Or UserList(userindex).flags.Privilegios > 2 Then Exit Sub
UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Cantidad
If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0

End Sub
Public Sub DoMineria(userindex As Integer, Mineral As Integer)
On Error GoTo errhandler
Dim MiObj As Obj
Dim Esfuerzo As Integer

If UserList(userindex).Clase = MINERO Then
    Esfuerzo = EsfuerzoExcavarMinero
Else: Esfuerzo = EsfuerzoExcavarGeneral
End If

If UserList(userindex).Stats.MinSta >= Esfuerzo Then
    Call QuitarSta(userindex, Esfuerzo)
    Call SendUserSTA(userindex)
Else
    Call SendData(ToIndex, userindex, 0, "9E")
    Exit Sub
End If

MiObj.OBJIndex = Mineral



If UserList(userindex).Clase = MINERO Then
    MiObj.Amount = Fix(4 + ((0.29 + 0.07 * Buleano(UserList(userindex).Recompensas(1) = 1 And UserList(userindex).Invent.HerramientaEqpObjIndex = PICO_EXPERTO)) * UserList(userindex).Stats.UserSkills(Mineria)))
Else: MiObj.Amount = 1
End If

If UserList(userindex).POS.Map <> 66 Then MiObj.Amount = MiObj.Amount * 5

If Not MeterItemEnInventario(userindex, MiObj) Then Call TirarItemAlPiso(UserList(userindex).POS, MiObj)
Call SubirSkill(userindex, Mineria, 5)
Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SOUND_MINERO)

Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub
Public Sub DoMeditar(userindex As Integer)

UserList(userindex).Counters.IdleCount = Timer

Dim Suerte As Integer
Dim Res As Integer
Dim Cant As Integer

If UserList(userindex).Stats.MinMAN >= UserList(userindex).Stats.MaxMAN Then
    Call SendData(ToIndex, userindex, 0, "D9")
    Call SendData(ToIndex, userindex, 0, "MEDOK")
    UserList(userindex).flags.Meditando = False
    UserList(userindex).Char.FX = 0
    UserList(userindex).Char.loops = 0
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
    Exit Sub
End If

If UserList(userindex).Stats.UserSkills(Meditar) <= 10 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 20 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 30 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 40 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 50 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 60 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 70 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 80 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 90 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(Meditar) <= 99 _
   And UserList(userindex).Stats.UserSkills(Meditar) >= 91 Then
                    Suerte = 8

ElseIf UserList(userindex).Stats.UserSkills(Meditar) = 100 Then

                    Suerte = 5
End If
Res = RandomNumber(1, Suerte)

If Res = 1 Then
    If UserList(userindex).Stats.MaxMAN > 0 Then Cant = Maximo(1, Porcentaje(UserList(userindex).Stats.MaxMAN, 3))
    Call AddtoVar(UserList(userindex).Stats.MinMAN, Cant, UserList(userindex).Stats.MaxMAN)
    Call SendData(ToIndex, userindex, 0, "MN" & Cant)
    Call SendUserMANA(userindex)
    Call SubirSkill(userindex, Meditar)
End If

End Sub
Public Sub InicioTrabajo(userindex As Integer, Trabajo As Long, TrabajoPos As WorldPos)


If Distancia(TrabajoPos, UserList(userindex).POS) > 2 Then
    Call SendData(ToIndex, userindex, 0, "DL")
    Exit Sub
End If


Select Case Trabajo
    
    

    Case Pesca
    
        
        If UserList(userindex).Invent.HerramientaEqpObjIndex <> OBJTYPE_CAÑA And UserList(userindex).Invent.HerramientaEqpObjIndex <> RED_PESCA Then
            Call SendData(ToIndex, userindex, 0, "%6")
            Exit Sub
        End If
        
        If MapData(UserList(userindex).POS.Map, TrabajoPos.x, TrabajoPos.Y).Agua = 0 Then
            Call SendData(ToIndex, userindex, 0, "6N")
            Exit Sub
        End If

    Case Talar
        
        If Trabajo = Talar Then
            If UserList(userindex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                Call SendData(ToIndex, userindex, 0, "%7")
                Exit Sub
            End If
        End If
        
        
        
        If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).trigger = 4 Then
            Call SendData(ToIndex, userindex, 0, "0W")
            Exit Sub
        End If

        If Not ObjData(MapData(TrabajoPos.Map, TrabajoPos.x, TrabajoPos.Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_ARBOLES Then
            Call SendData(ToIndex, userindex, 0, "2S")
            Exit Sub
        End If
                   
    Case Mineria
        
        If UserList(userindex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO And UserList(userindex).Invent.HerramientaEqpObjIndex <> PICO_EXPERTO Then
            Call SendData(ToIndex, userindex, 0, "%9")
            Exit Sub
        End If
        
        If Not ObjData(MapData(TrabajoPos.Map, TrabajoPos.x, TrabajoPos.Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_YACIMIENTO Then
            Call SendData(ToIndex, userindex, 0, "7N")
            Exit Sub
        End If

End Select


UserList(userindex).flags.Trabajando = Trabajo

UserList(userindex).TrabajoPos.x = TrabajoPos.x
UserList(userindex).TrabajoPos.Y = TrabajoPos.Y
Call SendData(ToIndex, userindex, 0, "%0")
Call SendData(ToIndex, userindex, 0, "MT")


End Sub
