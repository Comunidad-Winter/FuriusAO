Attribute VB_Name = "modHechizos"

    
Option Explicit
Sub NpcLanzaSpellSobreUser(NpcIndex As Integer, userindex As Integer, Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub

If Npclist(NpcIndex).ReyC > 0 Then
    If UserList(userindex).GuildInfo.GuildName <> "" Then
     If QuienConquista(Npclist(NpcIndex).ReyC) = UserList(userindex).GuildInfo.GuildName Then Exit Sub
    End If
End If

'If UserList(userindex).flags.Privilegios Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim Daño As Integer

If Hechizos(Spell).SubeHP = 1 Then
    If UserList(userindex).flags.Muerto = 1 Then Exit Sub
    
    Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + Daño
    If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call SendData(ToIndex, userindex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)
    Call SubirSkill(userindex, Resistencia)
ElseIf Hechizos(Spell).SubeHP = 2 Then
    Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)

    If Npclist(NpcIndex).MaestroUser = 0 Then Daño = Daño * (1 - UserList(userindex).Stats.UserSkills(Resistencia) / 200)

    If UserList(userindex).Invent.CascoEqpObjIndex Then
        Dim Obj As ObjData
       Obj = ObjData(UserList(userindex).Invent.CascoEqpObjIndex)
       If Obj.Gorro = 1 Then
       Dim absorbido As Integer
       absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
       absorbido = absorbido
       Daño = Maximo(1, Daño - absorbido)
       End If
    End If
    
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    
    If Not UserList(userindex).flags.Quest And UserList(userindex).flags.Privilegios = 0 Then
        UserList(userindex).Stats.MinHP = Maximo(0, UserList(userindex).Stats.MinHP - Daño)
        Call SendUserHP(userindex)
    End If
    
    Call SendData(ToIndex, userindex, 0, "%A" & Npclist(NpcIndex).Name & "," & Daño)
    Call SubirSkill(userindex, Resistencia)
    
    If UserList(userindex).Stats.MinHP = 0 Then Call UserDie(userindex)
    
End If
        
If Hechizos(Spell).Paraliza > 0 Then
     If UserList(userindex).flags.Paralizado = 0 Then
        If UserList(userindex).Clase = PIRATA And UserList(userindex).Recompensas(3) = 1 Then Exit Sub
        UserList(userindex).flags.Paralizado = 1
        UserList(userindex).Counters.Paralisis = Timer - 15 * (UserList(userindex).Clase = GUERRERO And UserList(userindex).Recompensas(3))
        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        Call SendData(ToIndex, userindex, 0, ("P9"))
        Call SendData(ToIndex, userindex, 0, "PU" & UserList(userindex).POS.x & "," & UserList(userindex).POS.Y)
     End If
End If

If Hechizos(Spell).Ceguera = 1 Then
    UserList(userindex).flags.Ceguera = 1
    UserList(userindex).Counters.Ceguera = Timer
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TX" & Hechizos(Spell).WAV & "," & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    Call SendData(ToIndex, userindex, 0, "CEGU")
    Call SendData(ToIndex, userindex, 0, "%B")
End If

If Hechizos(Spell).RemoverParalisis = 1 Then
     If Npclist(NpcIndex).flags.Paralizado Then
          Npclist(NpcIndex).flags.Paralizado = 0
          Npclist(NpcIndex).Contadores.Paralisis = 0
     End If
End If

End Sub
Function TieneHechizo(ByVal i As Integer, userindex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function
Sub AgregarHechizo(userindex As Integer, Slot As Byte)
Dim hIndex As Integer, j As Integer

hIndex = ObjData(UserList(userindex).Invent.Object(Slot).OBJIndex).HechizoIndex

If Not TieneHechizo(hIndex, userindex) Then
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
    Next
        
    If UserList(userindex).Stats.UserHechizos(j) Then
        Call SendData(ToIndex, userindex, 0, "%C")
    Else
        UserList(userindex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, userindex, CByte(j))
        
        Call QuitarUnItem(userindex, CByte(Slot))
    End If
Else
    Call SendData(ToIndex, userindex, 0, "%D")
End If

End Sub
Sub Aprenderhechizo(userindex As Integer, ByVal hechizoespecial As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = hechizoespecial

If Not TieneHechizo(hIndex, userindex) Then
    
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
    Next
        
    If UserList(userindex).Stats.UserHechizos(j) Then
        Call SendData(ToIndex, userindex, 0, "%C")
    Else
        UserList(userindex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, userindex, CByte(j))
        
    End If
Else
    Call SendData(ToIndex, userindex, 0, "%D")
End If

End Sub
Sub DecirPalabrasMagicas(ByVal S As String, userindex As Integer)
On Error Resume Next

Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "||" & vbCyan & "°" & S & "°" & UserList(userindex).Char.CharIndex)

End Sub
Function ManaHechizo(userindex As Integer, Hechizo As Integer) As Integer

If UserList(userindex).flags.Privilegios > 2 Or UserList(userindex).flags.Quest Then Exit Function

If UserList(userindex).Recompensas(3) = 1 And _
    ((UserList(userindex).Clase = Druida And Hechizo = 24) Or _
    (UserList(userindex).Clase = PALADIN And Hechizo = 10)) Then
    ManaHechizo = 250
ElseIf UserList(userindex).Clase = CLERIGO And UserList(userindex).Recompensas(3) = 2 And Hechizo = 11 Then
    ManaHechizo = 1100
Else: ManaHechizo = Hechizos(Hechizo).ManaRequerido
End If

End Function
Function PuedeLanzar(userindex As Integer, ByVal HechizoIndex As Integer) As Boolean
Dim wp2 As WorldPos

wp2.Map = UserList(userindex).flags.TargetMap
wp2.x = UserList(userindex).flags.TargetX
wp2.Y = UserList(userindex).flags.TargetY

If Not EnPantalla(UserList(userindex).POS, wp2, 1) Then Exit Function

If UserList(userindex).flags.Muerto Then
    Call SendData(ToIndex, userindex, 0, "MU")
    Exit Function
End If

If MapInfo(UserList(userindex).POS.Map).NoMagia Then
    Call SendData(ToIndex, userindex, 0, "/T")
    Exit Function
End If

If UserList(userindex).Stats.ELV < Hechizos(HechizoIndex).Nivel Then
    Call SendData(ToIndex, userindex, 0, "%%" & Hechizos(HechizoIndex).Nivel)
    Exit Function
End If

If UserList(userindex).Stats.UserSkills(Magia) < Hechizos(HechizoIndex).MinSkill Then
    Call SendData(ToIndex, userindex, 0, "%E")
    Exit Function
End If

If UserList(userindex).Stats.MinMAN < ManaHechizo(userindex, HechizoIndex) Then
    Call SendData(ToIndex, userindex, 0, "%F")
    Exit Function
End If

If UserList(userindex).Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
    Call SendData(ToIndex, userindex, 0, "9C")
    Exit Function
End If

PuedeLanzar = True

End Function

Sub InvocacionMercenario(userindex As Integer, NumObj As Integer)
Dim Masc As Integer

If UserList(userindex).NroMascotas >= MAXMASCOTAS Then Exit Sub

Dim ind As Integer, Index As Integer
Dim TargetPos As WorldPos

TargetPos.Map = UserList(userindex).POS.Map
TargetPos.x = UserList(userindex).POS.x + 1
TargetPos.Y = UserList(userindex).POS.Y

'For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(UserIndex).flags.Quest)
'    If UserList(UserIndex).MascotasIndex(j) Then
'        If Npclist(UserList(UserIndex).MascotasIndex(j)).Numero = Hechizos(H).NumNPC Then Masc = Masc + 1
'    End If
'Next

    If (UserList(userindex).NroMascotas < 3 Or UserList(userindex).flags.Quest) And UserList(userindex).NroMascotas < MAXMASCOTAS Then
        ind = SpawnNpc(NumObj, TargetPos, True, False)
        If ind < MAXNPCS Then
        
            UserList(userindex).NroMascotas = UserList(userindex).NroMascotas + 1
            
            Index = FreeMascotaIndex(userindex)
            
            UserList(userindex).MascotasIndex(Index) = ind
            UserList(userindex).MascotasType(Index) = Npclist(ind).Numero

            Npclist(ind).MaestroUser = userindex
'            Npclist(ind).Contadores.TiempoExistencia = Timer
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
    End If



End Sub

Sub HechizoInvocacion(userindex As Integer, B As Boolean)
Dim Masc As Integer

If Not MapInfo(UserList(userindex).POS.Map).Pk Then
    Call SendData(ToIndex, userindex, 0, "A&")
    Exit Sub
End If

If Not UserList(userindex).flags.Quest And UserList(userindex).NroMascotas >= 3 Then Exit Sub
If UserList(userindex).flags.EnReto Then Exit Sub
If UserList(userindex).NroMascotas >= MAXMASCOTAS Then Exit Sub

If UserList(userindex).Clase = Druida And UserList(userindex).Recompensas(3) <> 1 Then Exit Sub




Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos

TargetPos.Map = UserList(userindex).flags.TargetMap
TargetPos.x = UserList(userindex).flags.TargetX
TargetPos.Y = UserList(userindex).flags.TargetY

H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)

For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(j) Then
        If Npclist(UserList(userindex).MascotasIndex(j)).Numero = Hechizos(H).NumNPC Then Masc = Masc + 1
    End If
Next
'FuriusAO Se modifico el maximo de invocaciones del nigromante ahora solo saca 1 solo fuego y este tira inmo y desc
If (Hechizos(H).NumNPC = 103 And Masc >= 1 And Not UserList(userindex).flags.Quest) Or (Hechizos(H).NumNPC = 94 And Masc >= 1) Then
    Call SendData(ToIndex, userindex, 0, "||No puedes invocar más mascotas de este tipo." & FONTTYPE_FIGHT)
    Exit Sub
End If
'FuriusAO
For j = 1 To Hechizos(H).Cant
    If (UserList(userindex).NroMascotas < 3 Or UserList(userindex).flags.Quest) And UserList(userindex).NroMascotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNPC, TargetPos, True, False)
        If ind < MAXNPCS Then
        
            UserList(userindex).NroMascotas = UserList(userindex).NroMascotas + 1
                
            Index = FreeMascotaIndex(userindex)
            
            UserList(userindex).MascotasIndex(Index) = ind
            UserList(userindex).MascotasType(Index) = Npclist(ind).Numero
            
            If UserList(userindex).Clase = Druida And UserList(userindex).Recompensas(3) = 2 Then
                If Hechizos(H).NumNPC >= 92 And Hechizos(H).NumNPC <= 94 Then
                    Npclist(ind).Stats.MaxHP = Npclist(ind).Stats.MaxHP + 75
                    Npclist(ind).Stats.MinHP = Npclist(ind).Stats.MaxHP
                End If
            End If
            
            If Npclist(ind).Numero = 103 And UserList(userindex).Raza <> ELFO_OSCURO Then
                Npclist(ind).Stats.MaxHP = Npclist(ind).Stats.MaxHP - 200
                Npclist(ind).Stats.MinHP = Npclist(ind).Stats.MinHP - 200
            End If
            
            Npclist(ind).MaestroUser = userindex
            Npclist(ind).Contadores.TiempoExistencia = Timer
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
    Else: Exit For
    End If
Next

Call InfoHechizo(userindex)
B = True

End Sub
Sub HechizoTerrenoEstado(userindex As Integer, B As Boolean)
Dim PosCasteada As WorldPos
Dim TU As Integer
Dim H As Integer
Dim i As Integer

PosCasteada.x = UserList(userindex).flags.TargetX
PosCasteada.Y = UserList(userindex).flags.TargetY
PosCasteada.Map = UserList(userindex).flags.TargetMap

H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)

If Hechizos(H).Invisibilidad = 2 Then
    For i = 1 To MapInfo(UserList(userindex).POS.Map).NumUsers
        TU = MapInfo(UserList(userindex).POS.Map).userindex(i)
        If EnPantalla(PosCasteada, UserList(TU).POS, -1) And UserList(TU).flags.Invisible = 1 And UserList(TU).flags.AdminInvisible = 0 Then
        UserList(TU).flags.Oculto = 0
        UserList(TU).flags.Invisible = 0
        Call SendData(ToMap, 0, UserList(TU).POS.Map, ("V3" & UserList(TU).Char.CharIndex & ",0"))
        Call SendData(ToIndex, TU, 0, "V5")
        'Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(TU).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
        End If
    Next
    B = True
End If

Call InfoHechizo(userindex)

End Sub
Sub HandleHechizoTerreno(userindex As Integer, ByVal uh As Integer)
Dim B As Boolean

Select Case Hechizos(uh).Tipo
    Case uInvocacion
       Call HechizoInvocacion(userindex, B)
    Case uRadial
        Call HechizoTerrenoEstado(userindex, B)
End Select

If Hechizos(uh).TeleportX = 1 Then
'FIXIT: Declare 'Mapaf' and 'Xf' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim Mapaf, Xf, Yf As Integer
If UserList(userindex).flags.Portal = 0 Then
If MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).OBJInfo.OBJIndex <> 0 Then Exit Sub
If MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).trigger = 6 Then Exit Sub
If MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.Map <> 0 Then Exit Sub
If UserList(userindex).Counters.Pena > 0 Then Exit Sub
If UserList(userindex).POS.Map = Prision.Map Then Exit Sub

Dim TiemPoP As Integer

'CREAMOS EL PORTAL
'Dim ET As Obj
'ET.Amount = 1
'ET.OBJIndex = 862 'portal luminoso
'Mapaf = Hechizos(uh).TeleportXMap
'Xf = Hechizos(uh).TeleportXX
'Yf = Hechizos(uh).TeleportXY
'Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, ET, UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY)
'MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.Map = Mapaf
'MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.X = Xf
'MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.Y = Yf
'//// ACA TERMINAMOS DE CREARLO

'CREAMOS LA ANIMACION
Dim ET As Obj
ET.Amount = 1
ET.OBJIndex = 866 'PRE PORTAL LUMINOSO
Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, ET, UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY)
'EA



TiemPoP = 4 + 3 ' 4 segundos del portal + 3 DEL COSO
UserList(userindex).flags.Portal = TiemPoP
UserList(userindex).flags.PortalM = UserList(userindex).POS.Map
UserList(userindex).flags.PortalX = UserList(userindex).flags.TargetX
UserList(userindex).flags.PortalY = UserList(userindex).flags.TargetY
Call InfoHechizo(userindex)
B = True
Else
Call SendData(ToIndex, userindex, 0, "||No puedes lanzar mas de un portal a la vez" & FONTTYPE_INFO)
End If
End If


If Hechizos(uh).Materializa > 0 Then
Dim MAT As Obj

If MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).OBJInfo.OBJIndex <> 0 Then Exit Sub
If MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).TileExit.Map <> 0 Then Exit Sub
If MapData(UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY).Blocked = 1 Then Exit Sub
If MapInfo(UserList(userindex).POS.Map).Pk = False Then Exit Sub


MAT.Amount = Hechizos(uh).MaterializaCant
MAT.OBJIndex = Hechizos(uh).MaterializaObj
Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, MAT, UserList(userindex).POS.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY)
Call SendData(ToIndex, userindex, 0, "||Has materializado un objeto!!" & FONTTYPE_INFO)
Call InfoHechizo(userindex)
B = True
End If


If B Then
    Call SubirSkill(userindex, Magia)
    Call QuitarSta(userindex, Hechizos(uh).StaRequerido)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - ManaHechizo(userindex, uh)
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    Call SendUserMANASTA(userindex)
End If

End Sub
Sub HandleHechizoUsuario(userindex As Integer, ByVal uh As Integer)
Dim B As Boolean
Dim tempChr As Integer
'FIXIT: Declare 'TU' con un tipo de datos de enlace en tiempo de compilación               FixIT90210ae-R1672-R1B8ZE
Dim TU, tN As Integer

tempChr = UserList(userindex).flags.TargetUser

If UserList(tempChr).flags.Protegido = 1 Or UserList(tempChr).flags.Protegido = 2 Then Exit Sub

Select Case Hechizos(uh).Tipo
    Case uTerreno
       Call HechizoInvocacion(userindex, B)
    Case uEstado
       Call HechizoEstadoUsuario(userindex, B)
    Case uPropiedades
       Call HechizoPropUsuario(userindex, B)
End Select

If B Then
    Call SubirSkill(userindex, Magia)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - ManaHechizo(userindex, uh)
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    Call QuitarSta(userindex, Hechizos(uh).StaRequerido)
    Call SendUserMANASTA(userindex)
    Call SendUserHPSTA(UserList(userindex).flags.TargetUser)
    
    If Not UserList(UserList(userindex).flags.TargetUser).Name = UserList(userindex).Name Then Call QuitarInvisible(userindex)

    UserList(userindex).flags.TargetUser = 0
End If

End Sub
Sub HandleHechizoNPC(userindex As Integer, ByVal uh As Integer)
Dim B As Boolean

If Npclist(UserList(userindex).flags.TargetNpc).flags.NoMagia = 1 Then
    Call SendData(ToIndex, userindex, 0, "/U")
    Exit Sub
End If

If UserList(userindex).flags.Protegido > 0 Then
    Call SendData(ToIndex, userindex, 0, "||No podes atacar NPC's mientrás estas siendo protegido." & FONTTYPE_FIGHT)
    Exit Sub
End If

Select Case Hechizos(uh).Tipo
    Case uEstado
       Call HechizoEstadoNPC(UserList(userindex).flags.TargetNpc, uh, B, userindex)
    Case uPropiedades
       Call HechizoPropNPC(uh, UserList(userindex).flags.TargetNpc, userindex, B)
End Select

If B Then
   ' Call CheckPets(UserList(userindex).flags.TargetNpc, userindex)
    Call SubirSkill(userindex, Magia)
    UserList(userindex).flags.TargetNpc = 0
    Call QuitarSta(userindex, Hechizos(uh).StaRequerido)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - ManaHechizo(userindex, uh)
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    Call SendUserMANASTA(userindex)
    Call QuitarInvisible(userindex)
End If

End Sub
Sub LanzarHechizo(Index As Integer, userindex As Integer)
Dim uh As Integer
Dim exito As Boolean

If UserList(userindex).flags.Protegido = 1 Then
    Call SendData(ToIndex, userindex, 0, "||No podés tirar hechizos mientras estás siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Sub
ElseIf UserList(userindex).flags.Protegido = 2 Then
    Call SendData(ToIndex, userindex, 0, "||No podés tirar hechizos tan pronto al conectarte." & FONTTYPE_FIGHT)
    Exit Sub
End If

uh = UserList(userindex).Stats.UserHechizos(Index)

If (UserList(userindex).POS.Map = 148 Or UserList(userindex).POS.Map = 150) And (Hechizos(uh).Invoca > 0 Or Hechizos(uh).SubeHP = 2 Or Hechizos(uh).Invisibilidad = 1 Or Hechizos(uh).Paraliza > 0 Or Hechizos(uh).Estupidez = 1) Then
    Call SendData(ToIndex, userindex, 0, "||Una extraña energía te impide lanzar este hechizo..." & FONTTYPE_INFO)
    Exit Sub
End If

If TiempoTranscurrido(UserList(userindex).Counters.LastHechizo) < IntervaloUserPuedeCastear Then Exit Sub
If TiempoTranscurrido(UserList(userindex).Counters.LastGolpe - 0.5) < IntervaloUserPuedeGolpeHechi Then Exit Sub
UserList(userindex).Counters.LastHechizo = Timer
Call SendData(ToIndex, userindex, 0, "LH")
'FuriusAO Druida no usa mas baculo ni gorro
If Hechizos(uh).Baculo > 0 And UserList(userindex).Clase = MAGO Or UserList(userindex).Clase = NIGROMANTE Then
    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Baculo < Hechizos(uh).Baculo Then
        If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Baculo = 0 Then
            Call SendData(ToIndex, userindex, 0, "BN")
        Else: Call SendData(ToIndex, userindex, 0, "||Debes equiparte un báculo de mayor rango para lanzar este hechizo." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
End If
'FuriusAO
If PuedeLanzar(userindex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case uUsuarios
            If UserList(userindex).flags.TargetUser Then
                If UserList(UserList(userindex).flags.TargetUser).POS.Y - UserList(userindex).POS.Y >= 7 Then
                    Call SendData(ToIndex, userindex, 0, "||Estas demasiado lejos para lanzar ese hechizo." & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call HandleHechizoUsuario(userindex, uh)
            Else
                Call SendData(ToIndex, userindex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
                Call QuitarInvisible(userindex)
            End If
            
        Case uNPC
            If UserList(userindex).flags.TargetNpc Then
                Call HandleHechizoNPC(userindex, uh)
            Else
                Call SendData(ToIndex, userindex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If
            
        Case uUsuariosYnpc
            If UserList(userindex).flags.TargetUser Then
                If UserList(UserList(userindex).flags.TargetUser).POS.Y - UserList(userindex).POS.Y >= 7 Then
                    Call SendData(ToIndex, userindex, 0, "||Estas demasiado lejos para lanzar ese hechizo." & FONTTYPE_FIGHT)
                    Call QuitarInvisible(userindex)
                    Exit Sub
                End If
                Call HandleHechizoUsuario(userindex, uh)
            ElseIf UserList(userindex).flags.TargetNpc Then
                Call HandleHechizoNPC(userindex, uh)
            Else
                Call SendData(ToIndex, userindex, 0, "||Target invalido." & FONTTYPE_INFO)
                'Call QuitarInvisible(userindex)
            End If
            
        Case uTerreno
            Call HandleHechizoTerreno(userindex, uh)
        
        Case uArea
            Call HandleHechizoArea(userindex, uh)
        
    End Select
End If
                
End Sub
Sub HandleHechizoArea(userindex As Integer, ByVal uh As Integer)
On Error GoTo Error
Dim TargetPos As WorldPos
Dim X2 As Integer, Y2 As Integer
Dim UI As Integer
Dim B As Boolean

TargetPos.Map = UserList(userindex).flags.TargetMap
TargetPos.x = UserList(userindex).flags.TargetX
TargetPos.Y = UserList(userindex).flags.TargetY

For X2 = TargetPos.x - Hechizos(uh).RadioX To TargetPos.x + Hechizos(uh).RadioX
    For Y2 = TargetPos.Y - Hechizos(uh).RadioY To TargetPos.Y + Hechizos(uh).RadioY
        UI = MapData(TargetPos.Map, X2, Y2).userindex
        If UI > 0 Then
            UserList(userindex).flags.TargetUser = UI
            Select Case Hechizos(uh).Tipo
                Case uEstado
                    Call HechizoEstadoUsuario(userindex, B)
                Case uPropiedades
                    Call HechizoPropUsuario(userindex, B)
            End Select
        End If
    Next
Next

If B Then
    Call SubirSkill(userindex, Magia)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - ManaHechizo(userindex, uh)
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    Call QuitarSta(userindex, Hechizos(uh).StaRequerido)
    Call SendUserMANASTA(userindex)
    UserList(userindex).flags.TargetUser = 0
End If

Exit Sub
Error:
    Call LogError("Error en HandleHechizoArea")
End Sub
Public Function Amigos(userindex As Integer, UI As Integer) As Boolean

Amigos = (((UserList(userindex).Faccion.Bando = UserList(UI).Faccion.Bando) Or (EsNewbie(UI)) Or (EsNewbie(userindex)))) Or (UserList(userindex).POS.Map = 170) Or (UserList(userindex).POS.Map = 190) Or (UserList(userindex).Faccion.Bando = Neutral)

End Function
Sub HechizoEstadoUsuario(userindex As Integer, B As Boolean)
Dim H As Integer, TU As Integer, HechizoBueno As Boolean

H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
TU = UserList(userindex).flags.TargetUser

HechizoBueno = Hechizos(H).RemoverParalisis Or Hechizos(H).CuraVeneno Or Hechizos(H).Invisibilidad Or Hechizos(H).Revivir Or Hechizos(H).Flecha Or Hechizos(H).Estupidez = 2 Or Hechizos(H).Transforma

If HechizoBueno Then
    If Not Amigos(userindex, TU) Then
        Call SendData(ToIndex, userindex, 0, "2F")
        Exit Sub
    End If
Else
    If Not PuedeAtacar(userindex, TU) Then Exit Sub
    'If UserList(UserIndex).flags.Invisible Then Call BajarInvisible(UserIndex)
    Call UsuarioAtacadoPorUsuario(userindex, TU)
End If

 
    
If Hechizos(H).Envenena Then
    UserList(TU).flags.Envenenado = Hechizos(H).Envenena
    UserList(TU).flags.EstasEnvenenado = Timer
    UserList(TU).Counters.Veneno = Timer
    Call InfoHechizo(userindex)
    B = True
    Exit Sub
End If

If Hechizos(H).Maldicion = 1 Then
    UserList(TU).flags.Maldicion = 1
    'UserList(TU).Stats.MaxAGU = 0
    'UserList(TU).Stats.MaxAGU = 0
    'UserList(TU).Stats.MaxSta = 0
    Call InfoHechizo(userindex)
    B = True
    Exit Sub
End If

If Hechizos(H).Paraliza > 0 Then
     If UserList(TU).flags.Paralizado = 0 Then
        If (UserList(TU).Clase = MINERO And UserList(TU).Recompensas(2) = 1) Or (UserList(TU).Clase = MINERO And UserList(TU).POS.Map = 204) Then
             Call SendData(ToIndex, userindex, 0, "%&")
             Call SendData(ToIndex, userindex, 0, "PU" & UserList(userindex).POS.x & "," & UserList(userindex).POS.Y)
        'Or (UserList(TU).Clase = PIRATA And UserList(TU).Recompensas(3) = 1)
            Exit Sub
        End If
    
        UserList(TU).flags.QuienParalizo = userindex
        UserList(TU).flags.Paralizado = 1
        If (UserList(TU).Clase = PIRATA And UserList(TU).Recompensas(3) = 1) Then
        UserList(TU).Counters.Paralisis = Timer - (IntervaloParalizadoUsuario - 6)
        Else
        UserList(TU).Counters.Paralisis = Timer - 15 * Buleano(UserList(TU).Clase = GUERRERO And UserList(TU).Recompensas(3) = 2)
        End If
        Call SendData(ToIndex, TU, 0, "PU" & UserList(TU).POS.x & "," & UserList(TU).POS.Y)
        Call SendData(ToIndex, TU, 0, ("P9"))
        Call SumaPuntos(userindex, 1)
        Call InfoHechizo(userindex)
        B = True
        Exit Sub
    End If
End If
'leito nueevo a probar
'If Hechizo(H).Paraliza = 2 Then Call SendData(ToIndex, userindex, 0, "PU" & UserList(userindex).POS.X & "," & UserList(userindex).POS.Y)

If Hechizos(H).Ceguera = 1 Then
    UserList(TU).flags.Ceguera = 1
    UserList(TU).Counters.Ceguera = Timer
    Call SendData(ToIndex, TU, 0, "CEGU")
    Call InfoHechizo(userindex)
    B = True
    Exit Sub
End If

If Hechizos(H).Estupidez = 1 Then
If UserList(userindex).flags.EnReto Then Exit Sub
    UserList(TU).flags.Estupidez = 1
    UserList(TU).Counters.Estupidez = Timer
    Call SendData(ToIndex, TU, 0, "DUMB")
    Call InfoHechizo(userindex)
    B = True
    Exit Sub
End If

If Hechizos(H).Transforma = 1 Then
     If UserList(TU).flags.Transformado = 0 Then
        If UserList(TU).Stats.ELV > 39 And UserList(TU).Raza = ELFO And UserList(TU).Clase = Druida Then
            Call DoMetamorfosis(userindex)
        Else
            Call SendData(ToIndex, userindex, 0, "{E")
        End If
        Call InfoHechizo(userindex)
        B = True
        Exit Sub
    End If
End If

If Hechizos(H).Revivir = 1 Then
    If UserList(TU).flags.Muerto Then
    If UserList(userindex).flags.EnReto Then Exit Sub
        Call RevivirUsuario(userindex, TU, UserList(userindex).Clase = CLERIGO And UserList(userindex).Recompensas(3) = 2)
        Call InfoHechizo(userindex)
        B = True
        Exit Sub
    End If
End If

If UserList(TU).flags.Muerto Then
    Call SendData(ToIndex, userindex, 0, "8C")
    Exit Sub
End If

If Hechizos(H).Estupidez = 2 Then
    If UserList(TU).flags.Estupidez = 1 Then
        UserList(TU).flags.Estupidez = 0
        UserList(TU).Counters.Estupidez = 0
        Call SendData(ToIndex, TU, 0, "NESTUP")
        Call InfoHechizo(userindex)
        B = True
        Exit Sub
    End If
End If

If Hechizos(H).Flecha = 1 Then
    If TU <> userindex Then
        Call SendData(ToIndex, userindex, 0, "||Este hechizo solo puedes usarlo sobre ti mismo." & FONTTYPE_INFO)
        Exit Sub
    End If
    UserList(TU).flags.BonusFlecha = True
    UserList(TU).Counters.BonusFlecha = Timer
    Call InfoHechizo(userindex)
    B = True
    Exit Sub
End If

If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(TU).flags.Paralizado Then
    
    If UserList(userindex).flags.EnDM = True Then
    If userindex <> TU Then Exit Sub
    End If
    
    Call SendData(ToIndex, TU, 0, "P8")
        UserList(TU).flags.Paralizado = 0
        UserList(TU).flags.QuienParalizo = 0
        Call SumaPuntos(userindex, 2)
        Call InfoHechizo(userindex)
        B = True
        Exit Sub
    End If
End If

If Hechizos(H).Invisibilidad = 1 Then
    If UserList(TU).flags.Invisible Then Exit Sub
    
    If UserList(TU).Counters.uInvi > 0 Then
        Call SendData(ToIndex, TU, 0, "||Debes esperar " & UserList(TU).Counters.uInvi & " segundos para volver a estar invisible" & FONTTYPE_BLANCO)
        Exit Sub
    End If
    
    If MapData(UserList(TU).POS.Map, UserList(TU).POS.x, UserList(TU).POS.Y).trigger = 6 Then Exit Sub
    If UserList(TU).POS.Map = MAP_CTF Or UserList(TU).POS.Map = MAP_CTC Then Exit Sub
    UserList(TU).flags.Invisible = 1
    UserList(TU).Counters.Invisibilidad = Timer
    UserList(TU).Counters.uInvi = 6
    Call SendData(ToMap, 0, UserList(TU).POS.Map, ("V3" & UserList(TU).Char.CharIndex & ",1"))
    Call InfoHechizo(userindex)
    B = True
    Exit Sub
End If

If Hechizos(H).CuraVeneno = 1 Then
    If UserList(TU).flags.Envenenado = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(userindex)
        B = True
        Exit Sub
    Else
        Call SendData(ToIndex, userindex, 0, "||El usuario no está envenenado." & FONTTYPE_FIGHT)
        Exit Sub
    End If
End If

If Hechizos(H).RemoverMaldicion = 1 Then
    UserList(TU).flags.Maldicion = 0
    Call InfoHechizo(userindex)
    B = True
    Exit Sub
End If

If Hechizos(H).Bendicion = 1 Then
    UserList(TU).flags.Bendicion = 1
    Call InfoHechizo(userindex)
    B = True
    Exit Sub
End If

End Sub
Sub HechizoEstadoNPC(NpcIndex As Integer, ByVal hIndex As Integer, B As Boolean, userindex As Integer)

If Npclist(NpcIndex).Attackable = 0 Then Exit Sub

If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Invisible = 1
   B = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, userindex, 0, "NO")
        Exit Sub
   End If
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 1
   B = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 0
   B = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, userindex, 0, "NO")
        Exit Sub
   End If
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Maldicion = 1
   B = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Maldicion = 0
   B = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Bendicion = 1
   B = True
End If

If Hechizos(hIndex).Paraliza Then

If Npclist(NpcIndex).NoInmo = 1 Then
Call SendData(ToIndex, userindex, 0, "||El npc no se puede paralizar" & FONTTYPE_AZUL)
Exit Sub
End If


    If Npclist(NpcIndex).flags.QuienParalizo <> 0 And Npclist(NpcIndex).flags.QuienParalizo <> userindex Then Exit Sub
        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            Call InfoHechizo(userindex)
            Npclist(NpcIndex).flags.Paralizado = Hechizos(hIndex).Paraliza
            Npclist(NpcIndex).flags.QuienParalizo = userindex
            If Npclist(NpcIndex).flags.PocaParalisis = 1 Then
                Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado / 4
            Else: Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
            End If
            B = True
    Else: Call SendData(ToIndex, userindex, 0, "7D")
    End If
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
    If Npclist(NpcIndex).flags.QuienParalizo = userindex Or Npclist(NpcIndex).MaestroUser = userindex Then
       If Npclist(NpcIndex).flags.Paralizado Then
            Call InfoHechizo(userindex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            Npclist(NpcIndex).flags.QuienParalizo = 0
            B = True
       End If
    Else
        Call SendData(ToIndex, userindex, 0, "8D")
    End If
End If

End Sub
Sub VerNPCMuere(ByVal NpcIndex As Integer, ByVal Daño As Long, ByVal userindex As Integer)

If Npclist(NpcIndex).AutoCurar = 0 Then Npclist(NpcIndex).Stats.MinHP = Maximo(0, Npclist(NpcIndex).Stats.MinHP - Daño)

If Npclist(NpcIndex).Stats.MinHP <= 0 Then
    If Npclist(NpcIndex).flags.Snd3 Then Call SendData(ToNPCArea, NpcIndex, Npclist(NpcIndex).POS.Map, "TW" & Npclist(NpcIndex).flags.Snd3)
    
    If userindex Then
        If UserList(userindex).NroMascotas Then
            Dim T As Integer
            For T = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
                If UserList(userindex).MascotasIndex(T) Then
                    If Npclist(UserList(userindex).MascotasIndex(T)).TargetNpc = NpcIndex Then Call FollowAmo(UserList(userindex).MascotasIndex(T))
                End If
            Next
        End If
        Call AddtoVar(UserList(userindex).Stats.NPCsMuertos, 1, 32000)
        
        UserList(userindex).flags.TargetNpc = 0
        UserList(userindex).flags.TargetNpcTipo = 0
    End If
    
    Call MuereNpc(NpcIndex, userindex)
End If

End Sub
Sub ExperienciaPorGolpe(userindex As Integer, ByVal NpcIndex As Integer, Daño As Integer)
Dim ExpDada As Long

Daño = Minimo(Daño, Npclist(NpcIndex).Stats.MinHP)

ExpDada = Npclist(NpcIndex).GiveEXP * (Daño / Npclist(NpcIndex).Stats.MaxHP) / 2

If Daño >= Npclist(NpcIndex).Stats.MinHP Then ExpDada = ExpDada + Npclist(NpcIndex).GiveEXP / 2
If ModoQuest Then ExpDada = ExpDada / 2

If UserList(userindex).flags.Party = 0 Then
'If UserList(userindex).flags.Montado Then
'UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + (ExpDada / 2)
'ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinExp = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinExp + ExpDada / 100
'MascotaSubirExp (userindex)
'Else
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpDada
'End If

    If Daño >= Npclist(NpcIndex).Stats.MinHP Then
        Call SendData(ToIndex, userindex, 0, "EL" & ExpDada)
    Else: Call SendData(ToIndex, userindex, 0, "EX" & ExpDada)
    End If
    Call SendUserEXP(userindex)
    Call CheckUserLevel(userindex)
    Exit Sub

Else: Call RepartirExp(userindex, ExpDada, Daño >= Npclist(NpcIndex).Stats.MinHP)
End If

End Sub
Sub HechizoPropNPC(ByVal hIndex As Integer, NpcIndex As Integer, userindex As Integer, B As Boolean)
Dim Daño As Integer

If Npclist(NpcIndex).Attackable = 0 Then Exit Sub

If Hechizos(hIndex).SubeHP = 1 Then
    Daño = DañoHechizo(userindex, hIndex)
    
    Call InfoHechizo(userindex)
    Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, Daño, Npclist(NpcIndex).Stats.MaxHP)
    Call SendData(ToIndex, userindex, 0, "||" & vbYellow & "°+" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))
    Call SendData(ToIndex, userindex, 0, "CU" & Daño)
    B = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    Daño = DañoHechizo(userindex, hIndex)

    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Baculo = Hechizos(hIndex).Baculo Then Daño = 0.95 * Daño
    
    If UserList(userindex).flags.Montado = 1 Then
    Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHITMag, ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHITMag)
    End If
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, userindex, 0, "NO")
        Exit Sub
    End If
    
    If Npclist(NpcIndex).ReyC > 0 Then
        If UserList(userindex).GuildInfo.GuildName <> "" Then
            If QuienConquista(Npclist(NpcIndex).ReyC) = UserList(userindex).GuildInfo.GuildName Then Exit Sub
        End If
    End If

    

    If UserList(userindex).Faccion.Bando <> Neutral And Npclist(NpcIndex).MaestroUser Then
        If Not PuedeAtacarMascota(userindex, (Npclist(NpcIndex).MaestroUser)) Then Exit Sub
    End If
    
    If UserList(userindex).Faccion.Bando <> Neutral And UserList(userindex).Faccion.Bando = Npclist(NpcIndex).flags.Faccion Then
        Call SendData(ToIndex, userindex, 0, Mensajes(Npclist(NpcIndex).flags.Faccion, 19))
        Exit Sub
    End If
    
    Call InfoHechizo(userindex)
    B = True
    Call NpcAtacado(NpcIndex, userindex)
    
    If Npclist(NpcIndex).flags.Snd2 Then Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
     Call SendData(ToIndex, userindex, 0, "||" & vbYellow & "°-" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))
    Call SendData(ToIndex, userindex, 0, "X2" & Daño)
    
    Call ExperienciaPorGolpe(userindex, NpcIndex, Daño)
    
    Call VerNPCMuere(NpcIndex, Daño, userindex)
End If

End Sub
Sub InfoHechizo(userindex As Integer)
Dim H As Integer
H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)

Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, userindex)
Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & Hechizos(H).WAV)

If UserList(userindex).flags.TargetUser Then
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(UserList(userindex).flags.TargetUser).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
ElseIf UserList(userindex).flags.TargetNpc Then
    Call SendData(ToPCArea, userindex, Npclist(UserList(userindex).flags.TargetNpc).POS.Map, "CFX" & Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
End If

If UserList(userindex).flags.TargetUser Then
    If userindex <> UserList(userindex).flags.TargetUser Then
        Call SendData(ToIndex, userindex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(userindex).flags.TargetUser).Name & FONTTYPE_ATACO)
        Call SendData(ToIndex, UserList(userindex).flags.TargetUser, 0, "||" & UserList(userindex).Name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, userindex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
    End If
ElseIf UserList(userindex).flags.TargetNpc Then
    Call SendData(ToIndex, userindex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_ATACO)
End If
    
End Sub
Function TieneLaud(userindex) As Boolean
Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(i).OBJIndex = 469 Then
        'If UserList(userindex).Invent.Object(i).Equipped = 1 Then
        TieneLaud = True: Exit Function
    End If
Next
End Function
Function DañoHechizo(userindex As Integer, Hechizo As Integer) As Integer

DañoHechizo = RandomNumber(Hechizos(Hechizo).MinHP + 5 * Buleano(UserList(userindex).Clase = BARDO And UserList(userindex).Recompensas(3) = 2 And (Hechizo = 23 Or Hechizo = 25) And TieneLaud(userindex)) _
+ 10 * Buleano(UserList(userindex).Clase = NIGROMANTE And UserList(userindex).Recompensas(3) = 1) _
+ 20 * Buleano(UserList(userindex).Clase = CLERIGO And UserList(userindex).Recompensas(3) = 1 And Hechizo = 5) _
+ 10 * Buleano(UserList(userindex).Clase = MAGO And UserList(userindex).Recompensas(3) = 2 And Hechizo = 25), _
Hechizos(Hechizo).MaxHP + 5 * Buleano(UserList(userindex).Clase = BARDO And UserList(userindex).Recompensas(3) = 2 And (Hechizo = 23 Or Hechizo = 25) And TieneLaud(userindex)) _
+ 20 * Buleano(UserList(userindex).Clase = CLERIGO And UserList(userindex).Recompensas(3) = 1 And Hechizo = 5) _
+ 10 * Buleano(UserList(userindex).Clase = MAGO And UserList(userindex).Recompensas(3) = 1 And Hechizo = 23))

DañoHechizo = DañoHechizo + Porcentaje(DañoHechizo, 3 * UserList(userindex).Stats.ELV)

End Function
Sub HechizoPropUsuario(userindex As Integer, B As Boolean)
Dim H As Integer
Dim Daño As Integer
Dim tempChr As Integer
Dim reducido As Integer
Dim HechizoBueno As Boolean
Dim msg As String

H = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
tempChr = UserList(userindex).flags.TargetUser

HechizoBueno = Hechizos(H).SubeHam = 1 Or Hechizos(H).SubeSed = 1 Or Hechizos(H).SubeHP = 1 Or Hechizos(H).SubeAgilidad = 1 Or Hechizos(H).SubeFuerza = 1 Or Hechizos(H).SubeFuerza = 3 Or Hechizos(H).SubeMana = 1 Or Hechizos(H).SubeSta = 1

If HechizoBueno And Not Amigos(userindex, tempChr) Then
    Call SendData(ToIndex, userindex, 0, "2F")
    Exit Sub
ElseIf Not HechizoBueno Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    'If UserList(UserIndex).flags.Invisible Then Call BajarInvisible(UserIndex)
    Call UsuarioAtacadoPorUsuario(userindex, tempChr)
End If

If Hechizos(H).Revivir = 0 And UserList(tempChr).flags.Muerto Then Exit Sub

If Hechizos(H).SubeHam = 1 Then
    
    Daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    Call InfoHechizo(userindex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHam, Daño, UserList(tempChr).Stats.MaxHam)
    
    If userindex <> tempChr Then
        Call SendData(ToIndex, userindex, 0, "||Le has restaurado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(userindex).Name & " te ha restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, userindex, 0, "||Te has restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHyS(tempChr)
    B = True

ElseIf Hechizos(H).SubeHam = 2 Then
    Daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    UserList(tempChr).Stats.MinHam = Maximo(0, UserList(tempChr).Stats.MinHam - Daño)
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(ToIndex, userindex, 0, "||Le has quitado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(userindex).Name & " te ha quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, userindex, 0, "||Te has quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    If UserList(tempChr).Stats.MinHam = 0 Then UserList(tempChr).flags.Hambre = 1
    Call EnviarHyS(tempChr)
    B = True
End If


If Hechizos(H).SubeSed = 1 Then
    
    Call AddtoVar(UserList(tempChr).Stats.MinAGU, Daño, UserList(tempChr).Stats.MaxAGU)
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
      Call SendData(ToIndex, userindex, 0, "||Le has restaurado " & Daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
      Call SendData(ToIndex, tempChr, 0, "||" & UserList(userindex).Name & " te ha restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(ToIndex, userindex, 0, "||Te has restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    B = True

ElseIf Hechizos(H).SubeSed = 2 Then
    Daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    UserList(tempChr).Stats.MinAGU = Maximo(0, UserList(tempChr).Stats.MinAGU - Daño)
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(ToIndex, userindex, 0, "||Le has quitado " & Daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(userindex).Name & " te ha quitado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, userindex, 0, "||Te has quitado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU = 0 Then UserList(tempChr).flags.Sed = 1
    B = True
ElseIf Hechizos(H).SubeSed = 3 Then
    
    UserList(tempChr).Stats.MinAGU = 0
    UserList(tempChr).Stats.MinHam = 0
    UserList(tempChr).Stats.MinSta = 0
    UserList(tempChr).flags.Sed = 1
    UserList(tempChr).flags.Hambre = 1
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(ToIndex, userindex, 0, "S3" & UserList(tempChr).Name)
        Call SendData(ToIndex, tempChr, 0, "S4" & UserList(userindex).Name)
    Else
        Call SendData(ToIndex, userindex, 0, "S5")
    End If
    Call SendData(ToIndex, tempChr, 0, "2G")
    
    B = True
End If


If Hechizos(H).SubeAgilidad = 1 Then
    Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = Timer
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2, MAXATRIBUTOS))
    Call InfoHechizo(userindex)
    Call UpdateFuerzaYAg(tempChr)
    UserList(tempChr).flags.TomoPocion = True
    B = True

ElseIf Hechizos(H).SubeAgilidad = 2 Then
    UserList(tempChr).flags.TomoPocion = True
    Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = Timer
    Call RestVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, MINATRIBUTOS)
    Call InfoHechizo(userindex)
    Call UpdateFuerzaYAg(tempChr)
    B = True
ElseIf Hechizos(H).SubeAgilidad = 3 Then
    UserList(tempChr).flags.TomoPocion = True
    Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = Timer
    Call RestVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, MINATRIBUTOS)
    Call RestVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, MINATRIBUTOS)
    Call InfoHechizo(userindex)
    Call UpdateFuerzaYAg(tempChr)
    B = True
End If


If Hechizos(H).SubeFuerza = 1 Then
    Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = Timer
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(fuerza) * 2, MAXATRIBUTOS))
    Call InfoHechizo(userindex)
    Call UpdateFuerzaYAg(tempChr)
    UserList(tempChr).flags.TomoPocion = True
    B = True
ElseIf Hechizos(H).SubeFuerza = 2 Then
    UserList(tempChr).flags.TomoPocion = True
    Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = Timer
    Call RestVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, MINATRIBUTOS)
    Call InfoHechizo(userindex)
    Call UpdateFuerzaYAg(tempChr)
    B = True
ElseIf Hechizos(H).SubeFuerza = 3 Then
    Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = Timer
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(fuerza), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(fuerza) * 2, MAXATRIBUTOS))
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, Minimo(UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2, MAXATRIBUTOS))
    Call InfoHechizo(userindex)
    Call UpdateFuerzaYAg(tempChr)
    UserList(tempChr).flags.TomoPocion = True
    B = True
End If


If Hechizos(H).SubeHP = 1 Then
    If UserList(tempChr).flags.Muerto = 1 Then Exit Sub
    
    If UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP Then
        Call SendData(ToIndex, userindex, 0, "9D")
        Exit Sub
    End If
    Daño = DañoHechizo(userindex, H)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHP, Daño, UserList(tempChr).Stats.MaxHP)
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(ToIndex, userindex, 0, "R3" & Daño & "," & UserList(tempChr).Name)
        Call SendData(ToIndex, tempChr, 0, "R4" & UserList(userindex).Name & "," & Daño)
    Else
        Call SendData(ToIndex, userindex, 0, "R5" & Daño)
    End If
    B = True
ElseIf Hechizos(H).SubeHP = 2 Then
    Daño = DañoHechizo(userindex, H)
    
    If Hechizos(H).Baculo > 0 And (UserList(userindex).Clase = Druida Or UserList(userindex).Clase = MAGO Or UserList(userindex).Clase = NIGROMANTE) Then
        If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Baculo < Hechizos(H).Baculo Then
            Call SendData(ToIndex, userindex, 0, "BN")
            Exit Sub
        Else
            If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Baculo = Hechizos(H).Baculo Then Daño = 0.95 * Daño
        End If
    End If
    
    If UserList(tempChr).Invent.CascoEqpObjIndex Then
        Dim Obj As ObjData
        Obj = ObjData(UserList(tempChr).Invent.CascoEqpObjIndex)
        If Obj.Gorro = 1 Then Daño = Maximo(1, (1 - (RandomNumber(Obj.MinDef, Obj.MaxDef) / 100)) * Daño)
        Daño = Maximo(1, Daño)
    End If
    
    If UserList(userindex).flags.Montado = 1 Then
    Obj = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex)
    Daño = Daño + RandomNumber(Obj.MinHITMag, Obj.MaxHITMag)
    End If

    If UserList(tempChr).flags.Montado = 1 Then
    Obj = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex)
    Daño = Daño - RandomNumber(Obj.MinDefMag, Obj.MaxDefMag)
    If Daño < 0 Then Daño = 0
    End If

    If (UserList(tempChr).Clase = BARDO) And (UserList(tempChr).Recompensas(3) = 1) And TieneLaud(tempChr) Then
    Daño = Daño - (Daño / 22)
    End If


    If Not UserList(tempChr).flags.Quest Then UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - Daño
    Call InfoHechizo(userindex)
    
    Call SendData(ToIndex, userindex, 0, "6B" & Daño & "," & UserList(tempChr).Name)
    Call SendData(ToIndex, tempChr, 0, "7B" & Daño & "," & UserList(userindex).Name)
    'Call SendData(ToIndex, userindex, 0, "||" & vbYellow & "°-" & Daño & "°" & str(userindex))
    
    If UserList(tempChr).Stats.MinHP > 0 Then
        Call SubirSkill(tempChr, Resistencia)
    Else
        Call ContarMuerte(tempChr, userindex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, userindex)
    End If
    
    B = True
End If


If Hechizos(H).SubeMana = 1 Then
    Call AddtoVar(UserList(tempChr).Stats.MinMAN, Daño, UserList(tempChr).Stats.MaxMAN)
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(ToIndex, userindex, 0, "||Le has restaurado " & Daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(userindex).Name & " te ha restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, userindex, 0, "||Te has restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    B = True

ElseIf Hechizos(H).SubeMana = 2 Then

    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(ToIndex, userindex, 0, "||Le has quitado " & Daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(userindex).Name & " te ha quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, userindex, 0, "||Te has quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = Maximo(0, UserList(tempChr).Stats.MinMAN - Daño)
    B = True
    
End If


If Hechizos(H).SubeSta = 1 Then
    Call AddtoVar(UserList(tempChr).Stats.MinSta, Daño, UserList(tempChr).Stats.MaxSta)
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
         Call SendData(ToIndex, userindex, 0, "||Le has restaurado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
         Call SendData(ToIndex, tempChr, 0, "||" & UserList(userindex).Name & " te ha restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, userindex, 0, "||Te has restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    B = True
ElseIf Hechizos(H).SubeSta = 2 Then
    Call InfoHechizo(userindex)

    If userindex <> tempChr Then
        Call SendData(ToIndex, userindex, 0, "||Le has quitado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(userindex).Name & " te ha quitado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, userindex, 0, "||Te has quitado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    Call QuitarSta(tempChr, Daño)
    B = True
End If


End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, userindex As Integer, Slot As Byte)
Dim LoopC As Byte

If Not UpdateAll Then
    If UserList(userindex).Stats.UserHechizos(Slot) Then
        Call ChangeUserHechizo(userindex, Slot, UserList(userindex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(userindex, Slot, 0)
    End If
Else
    Call SendData(ToIndex, userindex, 0, "6H")
    For LoopC = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(LoopC) Then
            Call ChangeUserHechizo(userindex, LoopC, UserList(userindex).Stats.UserHechizos(LoopC))
        End If
    Next
End If

End Sub
Sub ChangeUserHechizo(userindex As Integer, Slot As Byte, ByVal Hechizo As Integer)

UserList(userindex).Stats.UserHechizos(Slot) = Hechizo

If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
    Call SendData(ToIndex, userindex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)
Else
    Call SendData(ToIndex, userindex, 0, "SHS" & Slot & "," & "0" & "," & "Nada")
End If

End Sub
Public Sub DesplazarHechizo(userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Byte)

If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then
    If CualHechizo = 1 Then
        Call SendData(ToIndex, userindex, 0, "%G")
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo - 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo - 1)
    End If
Else
    If CualHechizo = MAXUSERHECHIZOS Then
        Call SendData(ToIndex, userindex, 0, "%G")
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo + 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo + 1)
    End If
End If

Call UpdateUserHechizos(False, userindex, CualHechizo)

End Sub

