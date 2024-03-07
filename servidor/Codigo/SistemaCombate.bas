Attribute VB_Name = "SistemaCombate"

Option Explicit

Public Declare Function Minimo Lib "aolib.dll" (ByVal A As Long, ByVal B As Long) As Long
Public Declare Function Maximo Lib "aolib.dll" (ByVal A As Long, ByVal B As Long) As Long
Public Declare Function PoderAtaqueWresterling Lib "aolib.dll" (ByVal Skill As Byte, ByVal Agilidad As Integer, Clase As Byte, ByVal Nivel As Byte) As Integer
Public Declare Function SD Lib "aolib.dll" (ByVal N As Integer) As Integer
Public Declare Function SDM Lib "aolib.dll" (ByVal N As Integer) As Integer
Public Declare Function Complex Lib "aolib.dll" (ByVal N As Integer) As Integer
Public Declare Function RandomNumber Lib "aolib.dll" (ByVal MIN As Long, ByVal MAX As Long) As Long

Public Const EVASION = 1
Public Const CUERPOACUERPO = 2
Public Const CONARCOS = 3
Public Const EVAESCUDO = 4
Public Const DANOCUERPOACUERPO = 5
Public Const DANOCONARCOS = 6

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
Public Mods(1 To 6, 1 To NUMCLASES) As Single
Public Const MAXDISTANCIAARCO = 12
Public Sub CargarMods()
Dim i As Byte, j As Integer
Dim file As String

file = DatPath & "Mods.dat"

For i = 1 To NUMCLASES
    If Len(ListaClases(i)) > 0 Then
        For j = 1 To UBound(Mods, 1)
            Mods(j, i) = Int(GetVar(file, ListaClases(i), "Mod" & j)) / 100
        Next
    End If
Next

End Sub
Public Sub SaveMod(A As Integer, B As Integer)

Call WriteVar(DatPath & "Mods.dat", ListaClases(B), "Mod" & A, str(Mods(A, B) * 100))

End Sub

Public Function PoderAtaqueProyectil(userindex As Integer) As Integer

Select Case UserList(userindex).Stats.UserSkills(Proyectiles)
    Case Is < 31
        PoderAtaqueProyectil = UserList(userindex).Stats.UserSkills(Proyectiles) * Mods(CONARCOS, UserList(userindex).Clase)
    Case Is < 61
        PoderAtaqueProyectil = (UserList(userindex).Stats.UserSkills(Proyectiles) + UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(CONARCOS, UserList(userindex).Clase)
    Case Is < 91
        PoderAtaqueProyectil = (UserList(userindex).Stats.UserSkills(Proyectiles) + 2 * UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(CONARCOS, UserList(userindex).Clase)
    Case Else
        PoderAtaqueProyectil = (UserList(userindex).Stats.UserSkills(Proyectiles) + 3 * UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(CONARCOS, UserList(userindex).Clase)
End Select

PoderAtaqueProyectil = (PoderAtaqueProyectil + (2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0)))

End Function
Public Function PoderAtaqueArma(userindex As Integer) As Integer

Select Case UserList(userindex).Stats.UserSkills(Armas)
    Case Is < 31
        PoderAtaqueArma = UserList(userindex).Stats.UserSkills(Armas) * Mods(CUERPOACUERPO, UserList(userindex).Clase)
    Case Is < 61
        PoderAtaqueArma = (UserList(userindex).Stats.UserSkills(Armas) + UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(CUERPOACUERPO, UserList(userindex).Clase)
    Case Is < 91
        PoderAtaqueArma = (UserList(userindex).Stats.UserSkills(Armas) + 2 * UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(CUERPOACUERPO, UserList(userindex).Clase)
    Case Else
        PoderAtaqueArma = (UserList(userindex).Stats.UserSkills(Armas) + 3 * UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(CUERPOACUERPO, UserList(userindex).Clase)
End Select

PoderAtaqueArma = PoderAtaqueArma + 2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0)

End Function
'FIXIT: Declare 'PoderEvasionEscudo' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function PoderEvasionEscudo(userindex As Integer)

PoderEvasionEscudo = UserList(userindex).Stats.UserSkills(Defensa) * Mods(EVAESCUDO, UserList(userindex).Clase) / 2

End Function
Public Function PoderEvasion(userindex As Integer) As Integer

Select Case UserList(userindex).Stats.UserSkills(Tacticas)
    Case Is < 31
        PoderEvasion = UserList(userindex).Stats.UserSkills(Tacticas) * Mods(EVASION, UserList(userindex).Clase)
    Case Is < 61
        PoderEvasion = (UserList(userindex).Stats.UserSkills(Tacticas) + UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(EVASION, UserList(userindex).Clase)
    Case Is < 91
        PoderEvasion = (UserList(userindex).Stats.UserSkills(Tacticas) + 2 * UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(EVASION, UserList(userindex).Clase)
    Case Else
        PoderEvasion = (UserList(userindex).Stats.UserSkills(Tacticas) + 3 * UserList(userindex).Stats.UserAtributos(Agilidad)) * Mods(EVASION, UserList(userindex).Clase)
End Select

PoderEvasion = PoderEvasion + (2.5 * Maximo(UserList(userindex).Stats.ELV - 12, 0))

End Function
Public Function UserImpactoNpc(userindex As Integer, NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(userindex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma = 0 Then
    PoderAtaque = PoderAtaqueWresterling(UserList(userindex).Stats.UserSkills(Wresterling), UserList(userindex).Stats.UserAtributos(Agilidad), UserList(userindex).Clase, UserList(userindex).Stats.ELV) \ 4
ElseIf proyectil Then
    PoderAtaque = (1 + 0.05 * Buleano(UserList(userindex).Clase = ARQUERO And UserList(userindex).Recompensas(3) = 1) + 0.1 * Buleano(UserList(userindex).Recompensas(3) = 1 And (UserList(userindex).Clase = GUERRERO Or UserList(userindex).Clase = CAZADOR))) _
    * PoderAtaqueProyectil(userindex)
Else
    PoderAtaque = (1 + 0.05 * Buleano(UserList(userindex).Clase = PALADIN And UserList(userindex).Recompensas(3) = 2)) _
    * PoderAtaqueArma(userindex)
End If

ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma Then
       If proyectil Then
            Call SubirSkill(userindex, Proyectiles)
       Else: Call SubirSkill(userindex, Armas)
       End If
    Else
        Call SubirSkill(userindex, Wresterling)
    End If
End If


End Function
Public Function NpcImpacto(ByVal NpcIndex As Integer, userindex As Integer) As Boolean
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long

UserEvasion = (1 + 0.05 * Buleano(UserList(userindex).Recompensas(3) = 2 And (UserList(userindex).Clase = ARQUERO Or UserList(userindex).Clase = NIGROMANTE))) _
            * PoderEvasion(userindex)

If UserList(userindex).Invent.EscudoEqpObjIndex Then UserEvasion = UserEvasion + PoderEvasionEscudo(userindex)

ProbExito = Maximo(10, Minimo(90, 50 + ((Npclist(NpcIndex).PoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

If UserList(userindex).Invent.EscudoEqpObjIndex Then
Call SubirSkill(userindex, Defensa, 25)
   If Not NpcImpacto Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (UserList(userindex).Stats.UserSkills(Defensa) / (UserList(userindex).Stats.UserSkills(Defensa) + UserList(userindex).Stats.UserSkills(Tacticas)))))
      Rechazo = (RandomNumber(1, 34) <= ProbRechazo)
    
      If Rechazo Then
         Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_ESCUDO)
         Call SendData(ToIndex, userindex, 0, "7")
         Call SubirSkill(userindex, Defensa, 25)
      End If
   End If
End If

End Function
Public Function CalcularDaño(userindex As Integer, Optional ByVal Dragon As Boolean) As Long
Dim ModifClase As Single
Dim DañoUsuario As Long
Dim DañoArma As Long
Dim DañoMaxArma As Long
Dim Arma As ObjData

DañoUsuario = RandomNumber(UserList(userindex).Stats.MinHIT, UserList(userindex).Stats.MaxHIT)

If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
    ModifClase = Mods(DANOCUERPOACUERPO, UserList(userindex).Clase)
    CalcularDaño = Maximo(0, (UserList(userindex).Stats.UserAtributos(fuerza) - 15)) + DañoUsuario * ModifClase
    Exit Function
End If

Arma = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex)

DañoMaxArma = Arma.MaxHIT
        
If Arma.proyectil Then
    ModifClase = Mods(DANOCONARCOS, UserList(userindex).Clase)
    DañoArma = RandomNumber(Arma.MinHIT, DañoMaxArma) + RandomNumber(ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).MinHIT + 10 * Buleano(UserList(userindex).flags.BonusFlecha) + 5 * Buleano(UserList(userindex).Clase = ARQUERO And UserList(userindex).Recompensas(3) = 2), ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).MaxHIT + 15 * Buleano(UserList(userindex).flags.BonusFlecha) + 3 * Buleano(UserList(userindex).Clase = ARQUERO And UserList(userindex).Recompensas(3) = 2))
Else
    ModifClase = Mods(DANOCUERPOACUERPO, UserList(userindex).Clase)
    If Arma.SubTipo = MATADRAGONES And Not Dragon Then
        CalcularDaño = 1
        Exit Function
    Else
        DañoArma = RandomNumber(Arma.MinHIT, DañoMaxArma)
    End If
End If

CalcularDaño = (((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(userindex).Stats.UserAtributos(fuerza) - 15))) + DañoUsuario) * ModifClase)

End Function
Public Sub UserDañoNpc(userindex As Integer, ByVal NpcIndex As Integer)
Dim Muere As Boolean
Dim Daño As Long
Dim j As Integer

Daño = CalcularDaño(userindex, Npclist(NpcIndex).NPCtype = 6)

If Npclist(NpcIndex).ReyC > 0 Then
    If QuienConquista(Npclist(NpcIndex).ReyC) <> "" Then
        If val(Npclist(NpcIndex).Stats.MinHP) Mod 5 = 0 Or val(Npclist(NpcIndex).Stats.MinHP) Mod 2 = 0 Then
            Call SendData(ToAll, 0, 0, "||Fuerte " & Npclist(NpcIndex).ReyC & "> Está siendo atacado." & FONTTYPE_FUERTE)
        End If
    End If
End If

If UserList(userindex).flags.Navegando = 1 Then Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(userindex).Invent.BarcoObjIndex).MaxHIT)
If UserList(userindex).flags.Montado = 1 Then Daño = Daño + RandomNumber(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHIT, ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHIT)
Daño = Maximo(0, Daño - Npclist(NpcIndex).Stats.Def)

Call SendData(ToIndex, userindex, 0, "U2" & Daño)
Call SendData(ToIndex, userindex, 0, "||" & vbYellow & "°-" & Daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))

Call ExperienciaPorGolpe(userindex, NpcIndex, CInt(Daño))
If Daño >= Npclist(NpcIndex).Stats.MinHP Then Muere = True
Call VerNPCMuere(NpcIndex, Daño, userindex)

If Not Muere Then
    If PuedeApuñalar(userindex) Then
       Call DoApuñalar(userindex, NpcIndex, 0, CInt(Daño))
       Call SubirSkill(userindex, Apuñalar)
    End If
End If

End Sub
Public Sub NpcDaño(ByVal NpcIndex As Integer, userindex As Integer)
Dim Daño As Integer, lugar As Integer, absorbido As Integer, npcfile As String
Dim antdaño As Integer, defbarco As Integer
Dim Obj As ObjData
Dim Obj2 As ObjData

Daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antdaño = Daño

If UserList(userindex).flags.Navegando = 1 Then
    Obj = ObjData(UserList(userindex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If



If UserList(userindex).flags.Montado = 1 Then
    Obj = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

lugar = RandomNumber(1, 6)

Select Case lugar
  Case bCabeza
        
        If UserList(userindex).Invent.CascoEqpObjIndex Then
            Obj = ObjData(UserList(userindex).Invent.CascoEqpObjIndex)
            If Obj.Gorro = 0 Then absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
  Case Else
        
        If UserList(userindex).Invent.ArmourEqpObjIndex Then
           Obj = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If UserList(userindex).Invent.EscudoEqpObjIndex Then
           Obj2 = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex)
            absorbido = absorbido + RandomNumber(Obj2.MinDef, Obj2.MaxDef)
        End If
        
End Select

absorbido = absorbido + defbarco + 2 * Buleano(UserList(userindex).Clase = GUERRERO And UserList(userindex).Recompensas(2) = 2)

Daño = Maximo(1, Daño - absorbido)

Call SendData(ToIndex, userindex, 0, "N2" & lugar & "," & Daño)

If UserList(userindex).flags.Privilegios = 0 And Not UserList(userindex).flags.Quest Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - Daño

If UserList(userindex).Stats.MinHP <= 0 Then

    Call SendData(ToIndex, userindex, 0, "6")
    
   
    If Npclist(NpcIndex).MaestroUser Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
            Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
            Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
            Npclist(NpcIndex).flags.AttackedBy = 0
        End If
    End If
    
    Call UserDie(userindex)

End If

End Sub
Public Sub CheckPets(ByVal NpcIndex As Integer, userindex As Integer)
Dim j As Integer

For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(j) Then
       If UserList(userindex).MascotasIndex(j) <> NpcIndex Then
        If Npclist(UserList(userindex).MascotasIndex(j)).TargetNpc = 0 Then Npclist(UserList(userindex).MascotasIndex(j)).TargetNpc = NpcIndex
        Npclist(UserList(userindex).MascotasIndex(j)).Movement = NPC_ATACA_NPC
       End If
    End If
Next

End Sub
Public Sub AllFollowAmo(userindex As Integer)
Dim j As Integer

For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(j) Then
        Call FollowAmo(UserList(userindex).MascotasIndex(j))
    End If
Next

End Sub
Public Sub NpcAtacaUser(ByVal NpcIndex As Integer, userindex As Integer)

If Npclist(NpcIndex).AutoCurar = 1 Then Exit Sub
If Npclist(NpcIndex).Numero = 92 Then Exit Sub
If UserList(userindex).flags.Muerto Then Exit Sub

If Npclist(NpcIndex).ReyC > 0 Then
If UserList(userindex).GuildInfo.GuildName <> "" Then
If QuienConquista(Npclist(NpcIndex).ReyC) = UserList(userindex).GuildInfo.GuildName Then Exit Sub
End If
End If


If Npclist(NpcIndex).CanAttack = 1 Then
    Call CheckPets(NpcIndex, userindex)
    
    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = userindex
    
    If UserList(userindex).flags.AtacadoPorNpc = 0 And _
       UserList(userindex).flags.AtacadoPorUser = 0 Then UserList(userindex).flags.AtacadoPorNpc = NpcIndex
Else
    Exit Sub
End If

Npclist(NpcIndex).CanAttack = 0

If Npclist(NpcIndex).flags.Snd1 Then Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & Npclist(NpcIndex).flags.Snd1)
        
If NpcImpacto(NpcIndex, userindex) Then
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_IMPACTO)
    
    If UserList(userindex).flags.Navegando = 0 And Not UserList(userindex).flags.Meditando Then Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXSANGRE & "," & 0)

    Call NpcDaño(NpcIndex, userindex)

    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(userindex)
Else
    Call SendData(ToIndex, userindex, 0, "N1")
End If

Call SubirSkill(userindex, Tacticas)
Call SendUserHP(userindex)

End Sub
Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean

NpcImpactoNpc = (RandomNumber(1, 100) <= Maximo(10, Minimo(90, 50 + ((Npclist(Atacante).PoderAtaque - Npclist(Victima).PoderEvasion) * 0.4))))

End Function
Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
Dim Daño As Integer
Dim ANpc As Npc
ANpc = Npclist(Atacante)

Daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)

If ANpc.MaestroUser Then Call ExperienciaPorGolpe(ANpc.MaestroUser, Victima, Daño)
Call VerNPCMuere(Victima, Daño, ANpc.MaestroUser)

If Npclist(Victima).Stats.MinHP <= 0 Then
    Call RestoreOldMovement(Atacante)
    If ANpc.MaestroUser Then Call FollowAmo(Atacante)
End If

End Sub
Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer)

If Npclist(Atacante).CanAttack = 1 Then
    Npclist(Atacante).CanAttack = 0
    Npclist(Victima).TargetNpc = Atacante
Else: Exit Sub
End If

If Npclist(Atacante).flags.Snd1 Then Call SendData(ToNPCArea, Atacante, Npclist(Atacante).POS.Map, "TW" & Npclist(Atacante).flags.Snd1)

If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 Then
        Call SendData(ToNPCArea, Victima, Npclist(Victima).POS.Map, "TW" & Npclist(Victima).flags.Snd2)
    Else: Call SendData(ToNPCArea, Victima, Npclist(Victima).POS.Map, "TW" & SND_IMPACTO2)
    End If

    If Npclist(Atacante).MaestroUser Then
        Call SendData(ToNPCArea, Atacante, Npclist(Atacante).POS.Map, "TW" & SND_IMPACTO)
    Else: Call SendData(ToNPCArea, Victima, Npclist(Victima).POS.Map, "TW" & SND_IMPACTO)
    End If
    Call NpcDañoNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser Then
        Call SendData(ToNPCArea, Atacante, Npclist(Atacante).POS.Map, "TW" & SOUND_SWING)
    Else
        Call SendData(ToNPCArea, Victima, Npclist(Victima).POS.Map, "TW" & SOUND_SWING)
    End If
End If

End Sub
Public Sub UsuarioAtaca(userindex As Integer)
On Error Resume Next
If UserList(userindex).flags.Protegido = 1 Then
    Call SendData(ToIndex, userindex, 0, "||No podés atacar mientras estás siendo protegido por un GM." & FONTTYPE_INFO)
    Exit Sub
ElseIf UserList(userindex).flags.Protegido = 2 Then
    Call SendData(ToIndex, userindex, 0, "||No podés atacar tan pronto al conectarte." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(userindex).Clase = GUERRERO Then
If TiempoTranscurrido(UserList(userindex).Counters.LastGolpe - 0.3) < IntervaloUserPuedeAtacar Then Exit Sub
If TiempoTranscurrido(UserList(userindex).Counters.LastFlecha - 0.3) < IntervaloUserFlechas Then Exit Sub
Else
If TiempoTranscurrido(UserList(userindex).Counters.LastGolpe) < IntervaloUserPuedeAtacar Then Exit Sub
If TiempoTranscurrido(UserList(userindex).Counters.LastHechizo) < IntervaloUserPuedeHechiGolpe Then Exit Sub
If TiempoTranscurrido(UserList(userindex).Counters.LastFlecha) < IntervaloUserFlechas Then Exit Sub
End If

UserList(userindex).Counters.LastGolpe = Timer
Call SendData(ToIndex, userindex, 0, "LG")

If UserList(userindex).flags.Oculto Then
    If Not ((UserList(userindex).Clase = CAZADOR Or UserList(userindex).Clase = ARQUERO) And UserList(userindex).Invent.ArmourEqpObjIndex = 360) Then
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).flags.Invisible = 0
        Call SendData(ToMap, 0, UserList(userindex).POS.Map, ("V3" & UserList(userindex).Char.CharIndex & ",0"))
        Call SendData(ToIndex, userindex, 0, "V5")
    End If
End If

If UserList(userindex).Stats.MinSta >= 10 Then
    Call QuitarSta(userindex, RandomNumber(1, 10))
Else: Call SendData(ToIndex, userindex, 0, "9E")
    Exit Sub
End If

Dim AttackPos As WorldPos
AttackPos = UserList(userindex).POS
Call HeadtoPos(UserList(userindex).Char.Heading, AttackPos)

If AttackPos.x < XMinMapSize Or AttackPos.x > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "-" & UserList(userindex).Char.CharIndex)
    Exit Sub
End If

Dim Index As Integer
Index = MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).userindex

    If Index Then
            Call UsuarioAtacaUsuario(userindex, MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).userindex)
            Call SendUserSTA(userindex)
            Call SendUserHP(MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).userindex)
            Call QuitarInvisible(userindex)
            Exit Sub
    End If

If MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex Then

    If Npclist(MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex).Attackable Then
        
        If (Npclist(MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex).MaestroUser > 0 And _
           MapInfo(Npclist(MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex).POS.Map).Pk = False) And (UserList(userindex).POS.Map <> 190) Then
            Call SendData(ToIndex, userindex, 0, "0Z")
            Exit Sub
        End If
           
        Call UsuarioAtacaNpc(userindex, MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex)

    Else
        Call SendData(ToIndex, userindex, 0, "NO")
    End If
    
    Call SendUserSTA(userindex)
    
    Exit Sub


End If

Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "-" & UserList(userindex).Char.CharIndex)
Call SendUserSTA(userindex)

End Sub
Public Sub UsuarioAtacaNpc(userindex As Integer, ByVal NpcIndex As Integer)

'If Distancia(UserList(UserIndex).POS, Npclist(NpcIndex).POS) > MAXDISTANCIAARCO Then
'   Call SendData(ToIndex, UserIndex, 0, "3G")
'   Exit Sub
'End If

If (UserList(userindex).Faccion.Bando <> Neutral Or EsNewbie(userindex)) And Npclist(NpcIndex).MaestroUser Then
    If Not PuedeAtacarMascota(userindex, (Npclist(NpcIndex).MaestroUser)) Then Exit Sub
End If

If Npclist(NpcIndex).flags.Faccion <> Neutral Then
    If UserList(userindex).Faccion.Bando <> Neutral And UserList(userindex).Faccion.Bando = Npclist(NpcIndex).flags.Faccion Then
        Call SendData(ToIndex, userindex, 0, Mensajes(Npclist(NpcIndex).flags.Faccion, 19))
        Exit Sub
    ElseIf EsNewbie(userindex) Then
        Call SendData(ToIndex, userindex, 0, "%L")
        Exit Sub
    End If
End If

If UserList(userindex).flags.Protegido > 0 Then
    Call SendData(ToIndex, userindex, 0, "||No podes atacar NPC's mientrás estás siendo protegido." & FONTTYPE_FIGHT)
    Exit Sub
End If

If Npclist(NpcIndex).ReyC > 0 Then
    If UserList(userindex).GuildInfo.GuildName <> "" Then
        If QuienConquista(Npclist(NpcIndex).ReyC) = UserList(userindex).GuildInfo.GuildName Then Exit Sub
    End If
End If



Call NpcAtacado(NpcIndex, userindex)

If UserImpactoNpc(userindex, NpcIndex) Then
    If Npclist(NpcIndex).flags.Snd2 Then
        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "\" & UserList(userindex).Char.CharIndex & "," & Npclist(NpcIndex).flags.Snd2)
    Else
        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "?" & UserList(userindex).Char.CharIndex)
    End If
    Call UserDañoNpc(userindex, NpcIndex)
Else
     Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "-" & UserList(userindex).Char.CharIndex)
     Call SendData(ToIndex, userindex, 0, "U1")
End If

End Sub
Public Function TiempoTranscurrido(ByVal Desde As Single) As Single

TiempoTranscurrido = Timer - Desde

If TiempoTranscurrido < -5 Then
    TiempoTranscurrido = TiempoTranscurrido + 86400
ElseIf TiempoTranscurrido < 0 Then
    TiempoTranscurrido = 0
End If

End Function
Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
Dim ProbRechazo As Long
Dim Rechazo As Boolean
Dim ProbExito As Long
Dim PoderAtaque As Long
Dim UserPoderEvasion As Long
Dim proyectil As Boolean

If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex = 0 Then
    proyectil = False
Else: proyectil = ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil = 1
End If

UserPoderEvasion = (1 + 0.05 * Buleano(UserList(VictimaIndex).Recompensas(3) = 2 And (UserList(VictimaIndex).Clase = ARQUERO Or UserList(VictimaIndex).Clase = NIGROMANTE))) _
                    * PoderEvasion(VictimaIndex)

If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then UserPoderEvasion = UserPoderEvasion + PoderEvasionEscudo(VictimaIndex)


If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex Then
    If proyectil Then
        PoderAtaque = (1 + 0.05 * Buleano(UserList(AtacanteIndex).Clase = ARQUERO And UserList(AtacanteIndex).Recompensas(3) = 1) + 0.1 * Buleano(UserList(AtacanteIndex).Recompensas(3) = 1 And (UserList(AtacanteIndex).Clase = GUERRERO Or UserList(AtacanteIndex).Clase = CAZADOR))) _
        * PoderAtaqueProyectil(AtacanteIndex)
    Else
        PoderAtaque = (1 + 0.05 * Buleano(UserList(AtacanteIndex).Clase = PALADIN And UserList(AtacanteIndex).Recompensas(3) = 2)) _
        * PoderAtaqueArma(AtacanteIndex)
    End If
Else
    PoderAtaque = PoderAtaqueWresterling(UserList(AtacanteIndex).Stats.UserSkills(Wresterling), UserList(AtacanteIndex).Stats.UserAtributos(Agilidad), UserList(AtacanteIndex).Clase, UserList(AtacanteIndex).Stats.ELV)
End If

ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))

UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)


If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
    
    
    If Not UsuarioImpacto Then
      ProbRechazo = Maximo(10, Minimo(90, 100 * (UserList(VictimaIndex).Stats.UserSkills(Defensa) / (UserList(VictimaIndex).Stats.UserSkills(Defensa) + UserList(VictimaIndex).Stats.UserSkills(Tacticas)))))
      Rechazo = (RandomNumber(1, 50) <= ProbRechazo) 'ESTO CAMBIE
      If Rechazo Then
            
            Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).POS.Map, "&" & UserList(AtacanteIndex).Char.CharIndex)
            Call SendData(ToIndex, AtacanteIndex, 0, "8")
            Call SendData(ToIndex, VictimaIndex, 0, "7")
            Call SubirSkill(VictimaIndex, Defensa)
      End If
    End If
End If
    
If UsuarioImpacto Then
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex Then
        If Not proyectil Then
            Call SubirSkill(AtacanteIndex, Armas)
        Else: Call SubirSkill(AtacanteIndex, Proyectiles)
        End If
    Else
        Call SubirSkill(AtacanteIndex, Wresterling)
    End If
End If

End Function
Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

'If Distancia(UserList(AtacanteIndex).POS, UserList(VictimaIndex).POS) > MAXDISTANCIAARCO Then
'   Call SendData(ToIndex, AtacanteIndex, 0, "3G")
'   Exit Sub
'End If

Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
    'If UserList(AtacanteIndex).flags.Invisible Then Call BajarInvisible(AtacanteIndex)
    
    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).POS.Map, "TW" & "10")

    If UserList(VictimaIndex).flags.Navegando = 0 And Not UserList(VictimaIndex).flags.Meditando Then Call SendData(ToPCArea, VictimaIndex, UserList(VictimaIndex).POS.Map, "CFX" & UserList(VictimaIndex).Char.CharIndex & "," & FXSANGRE & "," & 0)
    
    Call UserDañoUser(AtacanteIndex, VictimaIndex)
    Call SumaPuntos(AtacanteIndex, 1)
Else
    Call SendData(ToPCArea, AtacanteIndex, UserList(AtacanteIndex).POS.Map, "-" & UserList(AtacanteIndex).Char.CharIndex)
    Call SendData(ToIndex, AtacanteIndex, 0, "U1")
    Call SendData(ToIndex, VictimaIndex, 0, "U3" & UserList(AtacanteIndex).Name)
End If

End Sub
Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
Dim Daño As Long, antdaño As Integer
Dim lugar As Integer, absorbido As Long
Dim defbarco As Integer
Dim Obj As ObjData
Dim Obj2 As ObjData
Dim j As Integer

Daño = CalcularDaño(AtacanteIndex)

antdaño = Daño



'MASCOTAS
If UserList(AtacanteIndex).flags.Montado = 1 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.MascotaEqpObjIndex)
     Daño = Daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
     Call SendData(ToIndex, AtacanteIndex, 0, "||Tu mascota ha pegado por " & Daño & FONTTYPE_BLANCO)
End If

If UserList(VictimaIndex).flags.Montado = 1 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.MascotaEqpObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If




If UserList(AtacanteIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(AtacanteIndex).Invent.BarcoObjIndex)
     Daño = Daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
End If

If UserList(VictimaIndex).flags.Navegando = 1 Then
     Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
     defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If






lugar = RandomNumber(1, 6)

Select Case lugar
  
  Case bCabeza
        If UserList(VictimaIndex).Invent.CascoEqpObjIndex Then
            If Not (UserList(AtacanteIndex).Clase = CAZADOR And UserList(AtacanteIndex).Recompensas(3) = 2) Then
                Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                  absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
            End If
        End If
        
  Case Else
        If UserList(VictimaIndex).Invent.ArmourEqpObjIndex Then
           Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
           Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
            absorbido = absorbido + RandomNumber(Obj2.MinDef, Obj2.MaxDef)
        End If
        
End Select

absorbido = absorbido + defbarco + 2 * Buleano(UserList(VictimaIndex).Clase = GUERRERO And UserList(VictimaIndex).Recompensas(2) = 2)
Daño = Maximo(1, Daño - absorbido)

Call SendData(ToIndex, AtacanteIndex, 0, "N5" & lugar & "," & Daño & "," & UserList(VictimaIndex).Name)
Call SendData(ToIndex, VictimaIndex, 0, "N4" & lugar & "," & Daño & "," & UserList(AtacanteIndex).Name)

If Not UserList(VictimaIndex).flags.Quest Then UserList(VictimaIndex).Stats.MinHP = UserList(VictimaIndex).Stats.MinHP - Daño

If UserList(AtacanteIndex).flags.Hambre = 0 And UserList(AtacanteIndex).flags.Sed = 0 Then
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex Then
        If ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).proyectil Then
            Call SubirSkill(AtacanteIndex, Proyectiles)
        Else: Call SubirSkill(AtacanteIndex, Armas)
        End If
    Else
        Call SubirSkill(AtacanteIndex, Wresterling)
    End If
    
    Call SubirSkill(AtacanteIndex, Tacticas)
    
    
    If PuedeApuñalar(AtacanteIndex) Then
        Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, Daño)
        Call SubirSkill(AtacanteIndex, Apuñalar)
    End If
End If

If UserList(VictimaIndex).Stats.MinHP <= 0 Then
     Call ContarMuerte(VictimaIndex, AtacanteIndex)
     Call SumaPuntos(AtacanteIndex, 4)
     

     For j = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(AtacanteIndex).flags.Quest)
        If UserList(AtacanteIndex).MascotasIndex(j) Then
            If Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = VictimaIndex Then Npclist(UserList(AtacanteIndex).MascotasIndex(j)).Target = 0
            Call FollowAmo(UserList(AtacanteIndex).MascotasIndex(j))
        End If
     Next

     Call ActStats(VictimaIndex, AtacanteIndex)
End If
        


Call CheckUserLevel(AtacanteIndex)


End Sub
Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

If UserList(AttackerIndex).POS.Map = 190 Or UserList(AttackerIndex).POS.Map = 170 Then Exit Sub

Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)

End Sub
Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(Maestro).flags.Quest)
    If UserList(Maestro).MascotasIndex(iCount) Then
        Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = victim
        Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = NPCDEFENSA
        Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next

End Sub
Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

If UserList(AttackerIndex).flags.Muerto Then
    Call SendData(ToIndex, AttackerIndex, 0, "MU")
    Exit Function
End If

If UserList(VictimIndex).flags.Muerto Then
    Call SendData(ToIndex, AttackerIndex, 0, "0X")
    Exit Function
End If

If AttackerIndex = VictimIndex Then
    If UserList(AttackerIndex).flags.Privilegios = 3 Then
        PuedeAtacar = True
        Exit Function
    Else
        Call SendData(ToIndex, AttackerIndex, 0, "%3")
        Exit Function
    End If
End If


  
    

If UserList(VictimIndex).flags.Party And UserList(AttackerIndex).PartyIndex = UserList(VictimIndex).PartyIndex Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a miembros de tu party." & FONTTYPE_FIGHT)
    Exit Function
End If

If MapData(UserList(VictimIndex).POS.Map, UserList(VictimIndex).POS.x, UserList(VictimIndex).POS.Y).trigger = 6 And MapData(UserList(AttackerIndex).POS.Map, UserList(AttackerIndex).POS.x, UserList(AttackerIndex).POS.Y).trigger Then
        PuedeAtacar = True
        Exit Function
End If


If UserList(VictimIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||El usuario está siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(VictimIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||El usuario está siendo protegido por conexión." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Privilegios >= 1 Then
    If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call SendData(ToIndex, AttackerIndex, 0, "E0")
    Exit Function
End If

If UserList(AttackerIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar mientrás estás siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(AttackerIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar tan pronto al conectarte." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).POS.Map <> 190 Or UserList(VictimIndex).POS.Map <> 170 Then
    If Not MapInfo(UserList(VictimIndex).POS.Map).Pk And Not TiempoTranscurrido(UserList(VictimIndex).Counters.LastRobo <= 10) Then
        Call SendData(ToIndex, AttackerIndex, 0, "7G")
        Exit Function
    End If
End If



If UserList(VictimIndex).POS.Map <> 170 Then
    If Not MapInfo(UserList(VictimIndex).POS.Map).Pk And Not TiempoTranscurrido(UserList(VictimIndex).Counters.LastRobo <= 10) Then
        Call SendData(ToIndex, AttackerIndex, 0, "7G")
        Exit Function
    End If
End If


If MapData(UserList(VictimIndex).POS.Map, UserList(VictimIndex).POS.x, UserList(VictimIndex).POS.Y).trigger = 4 Or MapData(UserList(AttackerIndex).POS.Map, UserList(AttackerIndex).POS.x, UserList(AttackerIndex).POS.Y).trigger = 4 Then
    Call SendData(ToIndex, AttackerIndex, 0, "8G")
    Exit Function
End If

If UserList(VictimIndex).POS.Map = 191 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a otros usuarios en el mapa de espera de torneo." & FONTTYPE_FIGHT)
    Exit Function
End If


If Not ModoQuest And Len(UserList(VictimIndex).GuildInfo.GuildName) > 0 And UserList(AttackerIndex).GuildInfo.GuildName = UserList(VictimIndex).GuildInfo.GuildName Then
    PuedeAtacar = True
    Exit Function
End If

  
        

 

If UserList(AttackerIndex).Faccion.Bando <> Neutral And UserList(AttackerIndex).Faccion.Bando = UserList(VictimIndex).Faccion.Bando Then
    If Len(UserList(AttackerIndex).GuildInfo.GuildName) = 0 Or Len(UserList(VictimIndex).GuildInfo.GuildName) = 0 Then
    '    Call SendData(ToIndex, AttackerIndex, 0, Mensajes(UserList(AttackerIndex).Faccion.Bando, 20))
        Exit Function
    End If
    If Not UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName) Then
    '    Call SendData(ToIndex, AttackerIndex, 0, Mensajes(UserList(AttackerIndex).Faccion.Bando, 20))
        Exit Function
    End If

    If ModoQuest Then
        Call SendData(ToIndex, AttackerIndex, 0, "||Durante una quest no puedes atacar a miembros de tu facción aunque pertenezcan a clanes enemigos." & FONTTYPE_FIGHT)
        Exit Function
    End If
End If

If EsNewbie(VictimIndex) And (EsNewbie(AttackerIndex) Or UserList(AttackerIndex).Faccion.Bando = Real) Then
    Call SendData(ToIndex, AttackerIndex, 0, "%1")
    Exit Function
End If

If EsNewbie(AttackerIndex) And UserList(VictimIndex).Faccion.Bando = Real Then
    Call SendData(ToIndex, AttackerIndex, 0, "%2")
    Exit Function
End If

PuedeAtacar = True

End Function

Public Function PuedeAtacarMascota(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

If AttackerIndex = VictimIndex Then
    PuedeAtacarMascota = True
    Exit Function
End If

If UserList(VictimIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar mascotas de usuarios protegidos por GMs." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(VictimIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar mascotas de usuarios protegidos por conexión." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Privilegios >= 1 Then Exit Function

If UserList(AttackerIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar mientrás estás siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(AttackerIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar tan pronto al conectarte." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).POS.Map <> 190 Or UserList(VictimIndex).POS.Map <> 170 Then
    If Not MapInfo(UserList(VictimIndex).POS.Map).Pk And Not TiempoTranscurrido(UserList(VictimIndex).Counters.LastRobo <= 10) Then
        Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar mascotas en zonas seguras." & FONTTYPE_FIGHT)
        Exit Function
    End If
End If



If MapData(UserList(VictimIndex).POS.Map, UserList(VictimIndex).POS.x, UserList(VictimIndex).POS.Y).trigger = 4 Or MapData(UserList(AttackerIndex).POS.Map, UserList(AttackerIndex).POS.x, UserList(AttackerIndex).POS.Y).trigger = 4 Then
    Call SendData(ToIndex, AttackerIndex, 0, "8G")
    Exit Function
End If

If Len(UserList(VictimIndex).GuildInfo.GuildName) > 0 And UserList(AttackerIndex).GuildInfo.GuildName = UserList(VictimIndex).GuildInfo.GuildName Then
    PuedeAtacarMascota = True
    Exit Function
End If

If UserList(AttackerIndex).Faccion.Bando <> Neutral And UserList(AttackerIndex).Faccion.Bando = UserList(VictimIndex).Faccion.Bando Then
    If Len(UserList(AttackerIndex).GuildInfo.GuildName) = 0 Or Len(UserList(VictimIndex).GuildInfo.GuildName) = 0 Then
        Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar mascotas de tu bando." & FONTTYPE_INFO)
        Exit Function
    End If
    If Not UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName) Then
        Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar mascotas de tu bando a menos que tu clan este en guerra con el del dueño." & FONTTYPE_INFO)
        Exit Function
    End If
End If

If EsNewbie(VictimIndex) And (EsNewbie(AttackerIndex) Or UserList(AttackerIndex).Faccion.Bando = Real) Then
    Call SendData(ToIndex, AttackerIndex, 0, "||Los miembros de la Alianza del Fúrius no pueden atacar mascotas de newbies." & FONTTYPE_INFO)
    Exit Function
End If

If EsNewbie(AttackerIndex) And UserList(VictimIndex).Faccion.Bando = Real Then
    Call SendData(ToIndex, AttackerIndex, 0, "||Los newbies no pueden atacar mascotas de la Alianza del Fúrius." & FONTTYPE_INFO)
    Exit Function
End If

If UserList(AttackerIndex).POS.Map = 190 Or UserList(AttackerIndex).POS.Map = 190 Then Exit Function


PuedeAtacarMascota = True

End Function
Public Function PuedeRobar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

If UserList(AttackerIndex).flags.Muerto Then
    Call SendData(ToIndex, AttackerIndex, 0, "MU")
    Exit Function
End If

If UserList(VictimIndex).flags.Muerto Then
    Call SendData(ToIndex, AttackerIndex, 0, "0X")
    Exit Function
End If

If AttackerIndex = VictimIndex Then
    Call SendData(ToIndex, AttackerIndex, 0, "%3")
    Exit Function
End If

If UserList(VictimIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||El usuario está siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(VictimIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||El usuario está siendo protegido por conexión." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Privilegios >= 1 Then
    If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call SendData(ToIndex, AttackerIndex, 0, "/F")
    Exit Function
End If

If UserList(AttackerIndex).flags.Protegido = 1 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar mientrás estás siendo protegido por un GM." & FONTTYPE_FIGHT)
    Exit Function
ElseIf UserList(AttackerIndex).flags.Protegido = 2 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No podes atacar tan pronto al conectarte." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).POS.Map <> 190 Or UserList(VictimIndex).POS.Map <> 170 Then
    If Not MapInfo(UserList(VictimIndex).POS.Map).Pk And Not TiempoTranscurrido(UserList(VictimIndex).Counters.LastRobo <= 10) Then
        Call SendData(ToIndex, AttackerIndex, 0, "/A")
        Exit Function
    End If
End If



If MapData(UserList(VictimIndex).POS.Map, UserList(VictimIndex).POS.x, UserList(VictimIndex).POS.Y).trigger = 4 Or MapData(UserList(AttackerIndex).POS.Map, UserList(AttackerIndex).POS.x, UserList(AttackerIndex).POS.Y).trigger = 4 Then
    Call SendData(ToIndex, AttackerIndex, 0, "/B")
    Exit Function
End If

If UserList(VictimIndex).POS.Map = 191 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a otros usuarios en el mapa de espera de torneo." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).flags.Party And UserList(AttackerIndex).PartyIndex = UserList(VictimIndex).PartyIndex Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar a miembros de tu party." & FONTTYPE_FIGHT)
    Exit Function
End If

If UserList(VictimIndex).Stats.MinSta < UserList(VictimIndex).Stats.MaxSta / 10 Then
    Call SendData(ToIndex, AttackerIndex, 0, "||No puedes atacar usuarios que tienen menos del 10% de su stamina total." & FONTTYPE_INFO)
    Exit Function
End If

If Len(UserList(VictimIndex).GuildInfo.GuildName) > 0 And UserList(AttackerIndex).GuildInfo.GuildName = UserList(VictimIndex).GuildInfo.GuildName Then
    PuedeRobar = True
    Exit Function
End If

If UserList(AttackerIndex).Faccion.Bando <> Neutral And UserList(AttackerIndex).Faccion.Bando = UserList(VictimIndex).Faccion.Bando Then
    If Len(UserList(AttackerIndex).GuildInfo.GuildName) = 0 Or Len(UserList(VictimIndex).GuildInfo.GuildName) = 0 Then
        Call SendData(ToIndex, AttackerIndex, 0, Mensajes(UserList(AttackerIndex).Faccion.Bando, 20))
        Exit Function
    End If
    If Not UserList(AttackerIndex).GuildRef.IsEnemy(UserList(VictimIndex).GuildInfo.GuildName) Then
        Call SendData(ToIndex, AttackerIndex, 0, Mensajes(UserList(AttackerIndex).Faccion.Bando, 20))
        Exit Function
    End If
End If

If EsNewbie(VictimIndex) And (EsNewbie(AttackerIndex) Or UserList(AttackerIndex).Faccion.Bando = Real) Then
    Call SendData(ToIndex, AttackerIndex, 0, "%1")
    Exit Function
End If

If EsNewbie(AttackerIndex) And UserList(VictimIndex).Faccion.Bando = Real Then
    Call SendData(ToIndex, AttackerIndex, 0, "%2")
    Exit Function
End If

PuedeRobar = True

Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)

End Function
