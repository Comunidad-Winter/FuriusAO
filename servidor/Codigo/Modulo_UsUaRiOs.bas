Attribute VB_Name = "UsUaRiOs"
Option Explicit
Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

If UserList(AttackerIndex).POS.Map <> 190 Or UserList(AttackerIndex).POS.Map <> 170 Then
    Dim DaExp As Integer
    DaExp = CInt(UserList(VictimIndex).Stats.ELV * RandomNumber(1, 4))
    Call AddtoVar(UserList(AttackerIndex).Stats.Exp, DaExp, MaxExp)
End If

Call SendData(ToIndex, AttackerIndex, 0, "1Q" & UserList(VictimIndex).Name)
Call SendData(ToIndex, AttackerIndex, 0, "EX" & DaExp)
Call SendData(ToIndex, VictimIndex, 0, "1R" & UserList(AttackerIndex).Name)

Call UserDie(VictimIndex)

If UserList(VictimIndex).Counters.Estupidez > 0 Then Call SendData(ToIndex, VictimIndex, 0, "NESTUP")

If UserList(VictimIndex).flags.Oferta > 0 Then
Dim UserU As Integer
UserU = NameIndex(UserList(VictimIndex).flags.Ofertador)
UserList(AttackerIndex).Stats.GLD = UserList(AttackerIndex).Stats.GLD + UserList(VictimIndex).flags.Oferta
'UserList(VictimIndex).Stats.GLD = UserList(VictimIndex).Stats.GLD + UserList(UserU).flags.Oferta
Call SendData(ToAll, AttackerIndex, 0, "||El caza recompensas " & UserList(AttackerIndex).Name & " ha matado a " & UserList(VictimIndex).Name & " y ha ganado " & UserList(VictimIndex).flags.Oferta & " por recompensa" & FONTTYPE_BLANCO)
'UserList(AttackerIndex).Stats.GLD = UserList(AttackerIndex).Stats.GLD + UserList(VictimIndex).flags.Oferta
Call SendUserStatsBox(AttackerIndex)
UserList(UserU).flags.Oferte = ""
UserList(VictimIndex).flags.Oferta = 0
UserList(VictimIndex).flags.Ofertador = ""
Call Ofertas.Quitar(UserList(VictimIndex).Name)
End If

If UserList(VictimIndex).flags.EnDM Then
Call ReincorporarDM(VictimIndex)
Call PagarDM(AttackerIndex)
UserList(VictimIndex).flags.DmMuertes = UserList(VictimIndex).flags.DmMuertes + 1
UserList(AttackerIndex).flags.DmKills = UserList(AttackerIndex).flags.DmKills + 1
Call SendData(ToIndex, VictimIndex, 0, "||" & DM_MMuerte & OroMuerte & " monedas de oro." & FONTTYPE_BLANCO)
Call SendData(ToIndex, AttackerIndex, 0, "||" & DM_MKill & OroKill & " monedas de oro!" & FONTTYPE_BLANCO)
End If


'TORNEO
If Torneo.Iniciado Then
If UserList(VictimIndex).flags.EnTorneo And UserList(AttackerIndex).flags.EnTorneo Then
Call PerdioRonda(VictimIndex, AttackerIndex)
End If
End If
'/torneo


If RetoEnCurso Or Reto2vs2EnCursO Then
If UserList(VictimIndex).flags.EnReto > 0 Then
    
    
If Reto2vs2EnCursO Then
    'If VictimIndex <> Pareja1.User1 And _
    'VictimIndex <> Pareja1.User2 And _
    'VictimIndex <> Pareja2.User1 And _
    'VictimIndex <> Pareja2.User2 Then
Select Case VictimIndex
    Case Pareja1.User1
        If UserList(Pareja1.User2).flags.Muerto = 1 Then
        'TERMINADO
        Call SendData(ToAll, 0, 0, "||Ring 2> " & UserList(Pareja2.User1).Name & " - " & UserList(Pareja2.User2).Name & " derrotaron a " & UserList(Pareja1.User1).Name & " - " & UserList(Pareja1.User2).Name & FONTTYPE_BLANCO)
        Call Pagar(2)
        Call DevolverParticipantes
        End If
    Exit Sub
    Case Pareja1.User2
        If UserList(Pareja1.User1).flags.Muerto = 1 Then
        'TERMINADO TAMBIEN
        Call SendData(ToAll, 0, 0, "||Ring 2> " & UserList(Pareja2.User1).Name & " - " & UserList(Pareja2.User2).Name & " derrotaron a " & UserList(Pareja1.User1).Name & " - " & UserList(Pareja1.User2).Name & FONTTYPE_BLANCO)
        Call Pagar(2)
        Call DevolverParticipantes
        End If
    Exit Sub
    Case Pareja2.User1
        If UserList(Pareja2.User2).flags.Muerto = 1 Then
        'TERMINADO
        Call SendData(ToAll, 0, 0, "||Ring 2> " & UserList(Pareja1.User1).Name & " - " & UserList(Pareja1.User2).Name & " derrotaron a " & UserList(Pareja2.User1).Name & " - " & UserList(Pareja2.User2).Name & FONTTYPE_BLANCO)
        Call Pagar(1)
        Call DevolverParticipantes
        End If
    Exit Sub
    Case Pareja2.User2
        If UserList(Pareja2.User1).flags.Muerto = 1 Then
        'TERMINADO TAMBIEN
        Call SendData(ToAll, 0, 0, "||Ring 2> " & UserList(Pareja1.User1).Name & " - " & UserList(Pareja1.User2).Name & " derrotaron a " & UserList(Pareja2.User1).Name & " - " & UserList(Pareja2.User2).Name & FONTTYPE_BLANCO)
        Call Pagar(1)
        Call DevolverParticipantes
        End If
    Exit Sub
    Case Else
    
End Select
   
End If

    
    If ItemsUvU = False Then
        Call WarpUserChar(AttackerIndex, 160, 51, 50, True)
    Else
        'TimerActiva
        frmMain.TimerRetos.Enabled = True
        TimerUVU = 6
        UVUname = AttackerIndex
        End If
        
        Call WarpUserChar(VictimIndex, 160, 50, 50, True)
        Call SendData(ToMap, 0, 160, "||Ring 1> " & UserList(AttackerIndex).Name & " derrotó a " & UserList(VictimIndex).Name & " en un reto." & FONTTYPE_BLANCO)
    
        UserList(AttackerIndex).flags.EnReto = 0
        UserList(VictimIndex).flags.EnReto = 0
        UserList(AttackerIndex).flags.RetadoPor = 0
        UserList(VictimIndex).flags.RetadoPor = 0
        UserList(AttackerIndex).flags.Retado = 0
        UserList(VictimIndex).flags.Retado = 0
        'RetoEnCurso = False
        UserList(AttackerIndex).flags.MatadasenR = UserList(AttackerIndex).flags.MatadasenR + 1
        UserList(VictimIndex).flags.PerdidasenR = UserList(VictimIndex).flags.PerdidasenR + 1
    End If
    
End If

End Sub
Sub RevivirUsuarioNPC(userindex As Integer)

UserList(userindex).flags.Muerto = 0
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP

Call DarCuerpoDesnudo(userindex)
Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
Call SendUserStatsBox(userindex)

End Sub
Sub RevivirUsuario(ByVal Resucitador As Integer, userindex As Integer, ByVal Lleno As Boolean)

UserList(Resucitador).Stats.MinSta = 0
UserList(Resucitador).Stats.MinAGU = 0
UserList(Resucitador).Stats.MinHam = 0
UserList(Resucitador).flags.Sed = 1
UserList(Resucitador).flags.Hambre = 1

UserList(userindex).flags.Muerto = 0

If Lleno Then
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN
    UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MaxHam
    UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MaxAGU
    UserList(userindex).flags.Sed = 0
    UserList(userindex).flags.Hambre = 0
Else
    UserList(userindex).Stats.MinHP = 1
    UserList(userindex).Stats.MinSta = 0
    UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinHam = 0
    UserList(userindex).Stats.MinAGU = 0
    UserList(userindex).flags.Sed = 1
    UserList(userindex).flags.Hambre = 1
End If

Call DarCuerpoDesnudo(userindex)
Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)

Call SendUserStatsBox(Resucitador)
Call EnviarHambreYsed(Resucitador)

Call SendUserStatsBox(userindex)
Call EnviarHambreYsed(userindex)

End Sub
Sub ReNombrar(userindex As Integer, NewNick As String)
If MySql = 0 Then
Kill CharPath & "/" & UCase$(UserList(userindex).Name) & ".CHR"
End If

Call SendData(ToIndex, userindex, 0, "||Has sido rebautizado como " & NewNick & "." & "~0~195~255~1~0")
Call SendData(ToAdmins, 0, 0, "||Servidor > El usuario " & UserList(userindex).Name & " ha sido rebautizado como " & NewNick & "." & FONTTYPE_BLANCO)
UserList(userindex).Name = NewNick
Call WarpUserChar(userindex, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y, False)

End Sub
Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

UserList(userindex).Char.Body = Body
UserList(userindex).Char.Head = Head
UserList(userindex).Char.Heading = Heading
UserList(userindex).Char.WeaponAnim = Arma
UserList(userindex).Char.ShieldAnim = Escudo
UserList(userindex).Char.CascoAnim = Casco

Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(userindex).Char.FX & "," & UserList(userindex).Char.loops & "," & Casco)

End Sub
Sub ChangeUserCharB(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, _
ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

On Error Resume Next

UserList(userindex).Char.Body = Body
UserList(userindex).Char.Head = Head
UserList(userindex).Char.Heading = Heading
UserList(userindex).Char.WeaponAnim = Arma
UserList(userindex).Char.ShieldAnim = Escudo
UserList(userindex).Char.CascoAnim = Casco

Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(userindex).Char.FX & "," & UserList(userindex).Char.loops & "," & Casco & "," & UserList(userindex).flags.Navegando)

End Sub
Sub ChangeUserCasco(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, _
ByVal Casco As Integer)

On Error Resume Next

If UserList(userindex).Char.CascoAnim <> Casco Then
UserList(userindex).Char.CascoAnim = Casco
Call SendData(sndRoute, sndIndex, sndMap, "7C" & UserList(userindex).Char.CharIndex & "," & Casco)
End If

End Sub
Sub ChangeUserEscudo(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, ByVal Escudo As Integer)
On Error Resume Next

If UserList(userindex).Char.ShieldAnim <> Escudo Then
    UserList(userindex).Char.ShieldAnim = Escudo
    Call SendData(sndRoute, sndIndex, sndMap, "6C" & UserList(userindex).Char.CharIndex & "," & Escudo)
End If

End Sub


Sub ChangeUserArma(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, _
ByVal Arma As Integer)

On Error Resume Next

If UserList(userindex).Char.WeaponAnim <> Arma Then
    UserList(userindex).Char.WeaponAnim = Arma
    Call SendData(sndRoute, sndIndex, sndMap, "5C" & UserList(userindex).Char.CharIndex & "," & Arma)
End If


End Sub


Sub ChangeUserHead(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, _
ByVal Head As Integer)

On Error Resume Next

If UserList(userindex).Char.Head <> Head Then
UserList(userindex).Char.Head = Head
Call SendData(sndRoute, sndIndex, sndMap, "4C" & UserList(userindex).Char.CharIndex & "," & Head)
End If

End Sub

Sub ChangeUserBody(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, _
ByVal Body As Integer)

On Error Resume Next
UserList(userindex).Char.Body = Body
Call SendData(sndRoute, sndIndex, sndMap, "3C" & UserList(userindex).Char.CharIndex & "," & Body)


End Sub
Sub ChangeUserHeading(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, userindex As Integer, _
ByVal Heading As Byte)
On Error Resume Next

UserList(userindex).Char.Heading = Heading
Call SendData(sndRoute, sndIndex, sndMap, "2C" & UserList(userindex).Char.CharIndex & "," & Heading)

End Sub
Sub EnviarSubirNivel(userindex As Integer, ByVal Puntos As Integer)

Call SendData(ToIndex, userindex, 0, "SUNI" & Puntos)

End Sub
Sub EnviarSkills(userindex As Integer)
Dim i As Integer
Dim cad As String

For i = 1 To NUMSKILLS
   cad = cad & UserList(userindex).Stats.UserSkills(i) & ","
Next

SendData ToIndex, userindex, 0, "SKILLS" & cad

End Sub
Sub EnviarFama(userindex As Integer)
Dim cad As String

cad = UserList(userindex).Faccion.Quests & ","
cad = cad & UserList(userindex).Faccion.Torneos & ","
    
If EsNewbie(userindex) Then
    cad = cad & UserList(userindex).Faccion.Matados(Caos) & ","
    cad = cad & UserList(userindex).Faccion.Matados(Neutral)
    
    Call SendData(ToIndex, userindex, 0, "FAMA3," & cad)
Else
    Select Case UserList(userindex).Faccion.Bando
        Case Neutral
            cad = cad & UserList(userindex).Faccion.BandoOriginal & ","
            cad = cad & UserList(userindex).Faccion.Matados(Real) & ","
            cad = cad & UserList(userindex).Faccion.Matados(Caos) & ","
            
        Case Real, Caos
            cad = cad & Titulo(userindex) & ","
            cad = cad & UserList(userindex).Faccion.Matados(Enemigo(UserList(userindex).Faccion.Bando)) & ","
            
    End Select
    cad = cad & UserList(userindex).Faccion.Matados(Neutral)
    Call SendData(ToIndex, userindex, 0, "FAMA" & UserList(userindex).Faccion.Bando & "," & cad)
End If

End Sub
Function GeneroLetras(Genero As Byte) As String

If Genero = 1 Then
    GeneroLetras = "Mujer"
Else
    GeneroLetras = "Hombre"
End If

End Function
Sub EnviarMiniSt(userindex As Integer)
Dim cad As String

cad = cad & UserList(userindex).Stats.VecesMurioUsuario & ","
cad = cad & UserList(userindex).Faccion.Matados(Caos) & ","
cad = cad & UserList(userindex).Stats.NPCsMuertos & ","
cad = cad & UserList(userindex).Faccion.Matados(Neutral) + UserList(userindex).Faccion.Matados(Real) + UserList(userindex).Faccion.Matados(Caos) & ","
cad = cad & ListaClases(UserList(userindex).Clase) & ","
cad = cad & ListaRazas(UserList(userindex).Raza) & ","
cad = cad & UserList(userindex).Faccion.Matados(Real) & ","

Call SendData(ToIndex, userindex, 0, "MIST" & cad)

End Sub
Sub EnviarAtrib(userindex As Integer)
Dim i As Integer
Dim cad As String

For i = 1 To NUMATRIBUTOS
  cad = cad & UserList(userindex).Stats.UserAtributos(i) & ","
Next

Call SendData(ToIndex, userindex, 0, "ATR" & cad)

End Sub
Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, userindex As Integer)

On Error GoTo ErrorHandler

CharList(UserList(userindex).Char.CharIndex) = 0

If UserList(userindex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).userindex = 0


Call SendData(ToMap, userindex, UserList(userindex).POS.Map, "BP" & UserList(userindex).Char.CharIndex)

UserList(userindex).Char.CharIndex = 0

NumChars = NumChars - 1

Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar; " & Err.Description)

End Sub
Sub UpdateUserChar(userindex As Integer)
On Error Resume Next
Dim bCr As Byte
Dim Info As String

If UserList(userindex).flags.Privilegios Then
    bCr = 1
ElseIf UserList(userindex).Faccion.Bando = Real Then
    bCr = 2
ElseIf UserList(userindex).Faccion.Bando = Caos Then
    bCr = 3
ElseIf EsNewbie(userindex) Then
    bCr = 4
Else: bCr = 5
End If

Info = "PW" & UserList(userindex).Char.CharIndex & "," & bCr & "," & UserList(userindex).Name

If Len(UserList(userindex).GuildInfo.GuildName) > 0 Then Info = Info & " <" & UserList(userindex).GuildInfo.GuildName & ">"

Call SendData(ToMap, userindex, UserList(userindex).POS.Map, (Info))

End Sub
Sub MakeUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, userindex As Integer, Map As Integer, x As Integer, Y As Integer)
On Error Resume Next
Dim CharIndex As Integer

If Not InMapBounds(x, Y) Then Exit Sub


If UserList(userindex).Char.CharIndex = 0 Then
    CharIndex = NextOpenCharIndex
    UserList(userindex).Char.CharIndex = CharIndex
    CharList(CharIndex) = userindex
End If


MapData(Map, x, Y).userindex = userindex


Dim klan$
klan$ = UserList(userindex).GuildInfo.GuildName
Dim bCr As Byte
If UserList(userindex).flags.Privilegios Then
    Select Case UserList(userindex).flags.Privilegios
    Case 3
        bCr = 1
    Case 2
        bCr = 6
    Case 1
        bCr = 7
    Case 4
        bCr = 8
    End Select
    
ElseIf UserList(userindex).flags.ConsejoCiuda Then
    bCr = 9
ElseIf UserList(userindex).flags.ConsejoCaoz Then
    bCr = 10
ElseIf UserList(userindex).Faccion.Bando = Real Then
    bCr = 2
ElseIf UserList(userindex).Faccion.Bando = Caos Then
    bCr = 3
ElseIf EsNewbie(userindex) Then
    bCr = 4
Else
    bCr = 5
End If

If Len(klan$) > 0 Then klan = " <" & klan$ & ">"

Call SendData(sndRoute, sndIndex, sndMap, ("CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & x & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).Name & klan$ & "," & bCr & "," & UserList(userindex).flags.Invisible))

If UserList(userindex).flags.Meditando Then
    UserList(userindex).Char.loops = LoopAdEternum
    If UserList(userindex).Stats.ELV < 15 Then
        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
        UserList(userindex).Char.FX = FXMEDITARCHICO
    ElseIf UserList(userindex).Stats.ELV < 30 Then
        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
        UserList(userindex).Char.FX = FXMEDITARMEDIANO
    Else
        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
        UserList(userindex).Char.FX = FXMEDITARGRANDE
    End If
End If

End Sub
Function Redondea(ByVal Number As Single) As Integer

If Number > Fix(Number) Then
    Redondea = Fix(Number) + 1
Else: Redondea = Number
End If

End Function
Sub CheckUserLevel(userindex As Integer)
On Error GoTo errhandler
Dim Pts As Integer
Dim SubeHit As Integer
Dim AumentoST As Integer
Dim AumentoMANA As Integer
Dim WasNewbie As Boolean

Do Until UserList(userindex).Stats.Exp < UserList(userindex).Stats.ELU
If UserList(userindex).Stats.ELV >= STAT_MAXELV Then
    UserList(userindex).Stats.Exp = 0
    UserList(userindex).Stats.ELU = 0
    Exit Sub
End If

WasNewbie = EsNewbie(userindex)

If UserList(userindex).Stats.Exp >= UserList(userindex).Stats.ELU Then

    If UserList(userindex).Stats.ELV >= 14 And ClaseBase(UserList(userindex).Clase) Then
        Call SendData(ToIndex, userindex, 0, "!6")
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU - 1
        Call SendUserEXP(userindex)
        Exit Sub
    End If
    
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SOUND_NIVEL)
    Call SendData(ToIndex, userindex, 0, "1S" & UserList(userindex).Stats.ELV + 1)
    
    If UserList(userindex).Stats.ELV = 1 Then
        Pts = 5
    Else
        Pts = 5
    End If
    
    UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts + Pts
    
    Call SendData(ToIndex, userindex, 0, "1T" & Pts)
       
    UserList(userindex).Stats.ELV = UserList(userindex).Stats.ELV + 1
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp - UserList(userindex).Stats.ELU
    UserList(userindex).Stats.ELU = ELUs(UserList(userindex).Stats.ELV)
    
    Dim AumentoHP As Integer
    Dim SubePromedio As Single
    
    SubePromedio = UserList(userindex).Stats.UserAtributos(Constitucion) / 2 - Resta(UserList(userindex).Clase)
    AumentoHP = RandomNumber(Fix(SubePromedio - 1), Redondea(SubePromedio + 1))
    SubeHit = AumentoHit(UserList(userindex).Clase)

    Select Case UserList(userindex).Clase
        Case CIUDADANO, TRABAJADOR, EXPERTO_MINERALES
            AumentoST = 15
            
        Case MINERO
            AumentoST = 15 + AdicionalSTMinero
            
        Case HERRERO
            AumentoST = 15
            
        Case EXPERTO_MADERA
            AumentoST = 15

        Case TALADOR
            AumentoST = 15 + AdicionalSTLeñador

        Case CARPINTERO
            AumentoST = 15
            
        Case PESCADOR
            AumentoST = 15 + AdicionalSTPescador
            
        Case SASTRE
            AumentoST = 15
            
        Case HECHICERO
            AumentoST = 15
            AumentoMANA = 2.2 * UserList(userindex).Stats.UserAtributos(Inteligencia)
            
        Case MAGO
            AumentoST = Maximo(5, 15 - AdicionalSTLadron / 2)
            Select Case UserList(userindex).Stats.MaxMAN
                Case Is < 2300
                    AumentoMANA = 3 * UserList(userindex).Stats.UserAtributos(Inteligencia)
                Case Is < 2500
                    AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(Inteligencia)
                Case Else
                    AumentoMANA = 1.5 * UserList(userindex).Stats.UserAtributos(Inteligencia)
            End Select
            
            If UserList(userindex).Stats.ELV > 45 Then AumentoMANA = 0
            
        Case NIGROMANTE
            AumentoST = Maximo(5, 15 - AdicionalSTLadron / 2)
            AumentoMANA = 2.2 * UserList(userindex).Stats.UserAtributos(Inteligencia)
            
        Case ORDEN_SAGRADA
            AumentoST = 15
            AumentoMANA = UserList(userindex).Stats.UserAtributos(Inteligencia)
            
        Case PALADIN
            AumentoST = 15
            AumentoMANA = UserList(userindex).Stats.UserAtributos(Inteligencia)
            
            If UserList(userindex).Stats.MaxHIT >= 99 Then SubeHit = 1
            
        Case CLERIGO
            AumentoST = 15
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(Inteligencia)

        Case NATURALISTA
            AumentoST = 15
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(Inteligencia)
            
        Case BARDO
            AumentoST = 15
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(Inteligencia)

        Case Druida
            AumentoST = 15
            AumentoMANA = 2.2 * UserList(userindex).Stats.UserAtributos(Inteligencia)

        Case SIGILOSO
            AumentoST = 15
            AumentoMANA = UserList(userindex).Stats.UserAtributos(Inteligencia)
            
        Case ASESINO
            AumentoST = 15
            AumentoMANA = UserList(userindex).Stats.UserAtributos(Inteligencia)

            If UserList(userindex).Stats.MaxHIT >= 99 Then SubeHit = 1
            
        Case CAZADOR
            AumentoST = 15
            AumentoMANA = UserList(userindex).Stats.UserAtributos(Inteligencia)

            If UserList(userindex).Stats.MaxHIT >= 99 Then SubeHit = 1
            
        Case SIN_MANA
            AumentoST = 15

        Case CABALLERO
            AumentoST = 15
            
        Case ARQUERO
            AumentoST = 15
         
            If UserList(userindex).Stats.MaxHIT >= 99 Then SubeHit = 2
            
        Case GUERRERO
            AumentoST = 15

            If UserList(userindex).Stats.MaxHIT >= 99 Then SubeHit = 2
           
        Case BANDIDO
            AumentoST = 15
            
        Case PIRATA
            AumentoST = 15

        Case LADRON
            AumentoST = 15
         
        Case Else
            AumentoST = 15 + AdicionalSTLadron
            
    End Select
       
    Call AddtoVar(UserList(userindex).Stats.MaxHP, AumentoHP, STAT_MAXHP)
    UserList(userindex).Stats.MaxSta = UserList(userindex).Stats.MaxSta + AumentoST
    
    Call AddtoVar(UserList(userindex).Stats.MaxMAN, AumentoMANA, 2200 + 800 * Buleano(UserList(userindex).Clase And UserList(userindex).Recompensas(2) = 2))
    UserList(userindex).Stats.MaxHIT = UserList(userindex).Stats.MaxHIT + SubeHit
    UserList(userindex).Stats.MinHIT = UserList(userindex).Stats.MinHIT + SubeHit
    
    Call SendData(ToIndex, userindex, 0, "1U" & AumentoHP & "," & AumentoST & "," & AumentoMANA & "," & SubeHit)
    
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, Pts)
   
    Call SendUserStatsBox(userindex)
    
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & ",19,1")
    
    If Not EsNewbie(userindex) And WasNewbie Then
        If UserList(userindex).POS.Map = 37 Or UserList(userindex).POS.Map = 49 Then
            Call WarpUserChar(userindex, 1, 50, 50, True)
        Else
            Call UpdateUserChar(userindex)
        End If
        Call QuitarNewbieObj(userindex)
        Call SendData(ToIndex, userindex, 0, "SUFA1")
    End If
    
    Call CheckUserLevel(userindex)
    
Else

    Call SendUserEXP(userindex)
    
End If

    
If PuedeSubirClase(userindex) Then Call SendData(ToIndex, userindex, 0, "SUCL1")
If PuedeRecompensa(userindex) Then Call SendData(ToIndex, userindex, 0, "SURE1")

Loop

Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub
Function PuedeRecompensa(userindex As Integer) As Byte

If UserList(userindex).Clase = SASTRE Then Exit Function

If UserList(userindex).Recompensas(1) = 0 And UserList(userindex).Stats.ELV >= 18 Then
    PuedeRecompensa = 1
    Exit Function
End If

If UserList(userindex).Clase = TALADOR Or UserList(userindex).Clase = PESCADOR Then Exit Function

If UserList(userindex).Stats.ELV >= 25 And UserList(userindex).Recompensas(2) = 0 Then
    PuedeRecompensa = 2
    Exit Function
End If
    
If UserList(userindex).Clase = CARPINTERO Then Exit Function

If UserList(userindex).Recompensas(3) = 0 And _
    (UserList(userindex).Stats.ELV >= 34 Or _
    (ClaseTrabajadora(UserList(userindex).Clase) And UserList(userindex).Stats.ELV >= 32) Or _
    ((UserList(userindex).Clase = PIRATA Or UserList(userindex).Clase = LADRON) And UserList(userindex).Stats.ELV >= 30)) Then
    PuedeRecompensa = 3
    Exit Function
End If

End Function
Function PuedeFaccion(userindex As Integer) As Boolean

PuedeFaccion = Not EsNewbie(userindex) And UserList(userindex).Faccion.BandoOriginal = Neutral And Len(UserList(userindex).GuildInfo.GuildName) = 0 And UserList(userindex).flags.Privilegios = 0

End Function
Function PuedeSubirClase(userindex As Integer) As Boolean

PuedeSubirClase = (UserList(userindex).Stats.ELV >= 3 And UserList(userindex).Clase = CIUDADANO) Or _
                (UserList(userindex).Stats.ELV >= 6 And (UserList(userindex).Clase = LUCHADOR Or UserList(userindex).Clase = TRABAJADOR)) Or _
                (UserList(userindex).Stats.ELV >= 9 And (UserList(userindex).Clase = EXPERTO_MINERALES Or UserList(userindex).Clase = EXPERTO_MADERA Or UserList(userindex).Clase = CON_MANA Or UserList(userindex).Clase = SIN_MANA)) Or _
                (UserList(userindex).Stats.ELV >= 12 And (UserList(userindex).Clase = CABALLERO Or UserList(userindex).Clase = BANDIDO Or UserList(userindex).Clase = HECHICERO Or UserList(userindex).Clase = NATURALISTA Or UserList(userindex).Clase = ORDEN_SAGRADA Or UserList(userindex).Clase = SIGILOSO))

End Function
Function PuedeAtravesarAgua(userindex As Integer) As Boolean

PuedeAtravesarAgua = UserList(userindex).flags.Navegando = 1 'Or UserList(userindex).flags.Privilegios > 0

End Function
Private Sub EnviaNuevaPosUsuarioPj(userindex As Integer, ByVal Quien As Integer)

Call SendData(ToIndex, userindex, 0, ("LP" & UserList(Quien).Char.CharIndex & "," & UserList(Quien).POS.x & "," & UserList(Quien).POS.Y & "," & UserList(Quien).Char.Heading))

End Sub
Private Sub EnviaNuevaPosNPC(userindex As Integer, NpcIndex As Integer)

Call SendData(ToIndex, userindex, 0, ("LP" & Npclist(NpcIndex).Char.CharIndex & "," & Npclist(NpcIndex).POS.x & "," & Npclist(NpcIndex).POS.Y & "," & Npclist(NpcIndex).Char.Heading))

End Sub
Sub CalcularValores(userindex As Integer)
Dim SubePromedio As Single
Dim HPReal As Integer
Dim HitReal As Integer
Dim i As Integer

HPReal = 15 + RandomNumber(1, UserList(userindex).Stats.UserAtributos(Constitucion) \ 3)
HitReal = AumentoHit(UserList(userindex).Clase) * UserList(userindex).Stats.ELV
SubePromedio = UserList(userindex).Stats.UserAtributos(Constitucion) / 2 - Resta(UserList(userindex).Clase)

For i = 1 To UserList(userindex).Stats.ELV - 1
    HPReal = HPReal + RandomNumber(Redondea(SubePromedio - 2), Fix(SubePromedio + 2))
Next

Call CalcularMana(userindex)

UserList(userindex).Stats.MinHIT = HitReal
UserList(userindex).Stats.MaxHIT = HitReal + 1
    
UserList(userindex).Stats.MinHP = Minimo(UserList(userindex).Stats.MinHP, HPReal)
UserList(userindex).Stats.MaxHP = HPReal
Call SendUserStatsBox(userindex)

End Sub
Sub CalcularMana(userindex As Integer)
Dim ManaReal As Integer

Select Case (UserList(userindex).Clase)
    Case HECHICERO
        ManaReal = 100 + 2.2 * (UserList(userindex).Stats.UserAtributos(Inteligencia) * (UserList(userindex).Stats.ELV - 1))
    
    Case MAGO
        ManaReal = 100 + 3 * (UserList(userindex).Stats.UserAtributos(Inteligencia) * (UserList(userindex).Stats.ELV - 1))
        
    Case ORDEN_SAGRADA
        ManaReal = UserList(userindex).Stats.UserAtributos(Inteligencia) * (UserList(userindex).Stats.ELV - 1)
    
    Case CLERIGO
        ManaReal = 50 + 2 * UserList(userindex).Stats.UserAtributos(Inteligencia) * (UserList(userindex).Stats.ELV - 1)

    Case NATURALISTA
        ManaReal = 50 + 2 * UserList(userindex).Stats.UserAtributos(Inteligencia) * (UserList(userindex).Stats.ELV - 1)

    Case Druida
        ManaReal = 50 + 2.1 * UserList(userindex).Stats.UserAtributos(Inteligencia) * (UserList(userindex).Stats.ELV - 1)
        
    Case SIGILOSO
        ManaReal = 50 + UserList(userindex).Stats.UserAtributos(Inteligencia) * (UserList(userindex).Stats.ELV - 1)
End Select

If ManaReal Then
    UserList(userindex).Stats.MinMAN = Minimo(UserList(userindex).Stats.MinMAN, ManaReal)
    UserList(userindex).Stats.MaxMAN = ManaReal
End If

End Sub
Private Sub EnviaGenteEnNuevoRango(userindex As Integer, ByVal nHeading As Byte)
Dim x As Integer, Y As Integer
Dim M As Integer

M = UserList(userindex).POS.Map

Select Case nHeading

Case NORTH, SOUTH

    If nHeading = NORTH Then
        Y = UserList(userindex).POS.Y - MinYBorder - 3
    Else
        Y = UserList(userindex).POS.Y + MinYBorder + 3
    End If
    For x = UserList(userindex).POS.x - MinXBorder - 2 To UserList(userindex).POS.x + MinXBorder + 2
        If MapData(M, x, Y).userindex Then
            Call EnviaNuevaPosUsuarioPj(userindex, MapData(M, x, Y).userindex)
        ElseIf MapData(M, x, Y).NpcIndex Then
            Call EnviaNuevaPosNPC(userindex, MapData(M, x, Y).NpcIndex)
        End If
    Next
Case EAST, WEST

    If nHeading = EAST Then
        x = UserList(userindex).POS.x + MinXBorder + 3
    Else
        x = UserList(userindex).POS.x - MinXBorder - 3
    End If
    For Y = UserList(userindex).POS.Y - MinYBorder - 2 To UserList(userindex).POS.Y + MinYBorder + 2
        If MapData(M, x, Y).userindex Then
            Call EnviaNuevaPosUsuarioPj(userindex, MapData(M, x, Y).userindex)
        ElseIf MapData(M, x, Y).NpcIndex Then
            Call EnviaNuevaPosNPC(userindex, MapData(M, x, Y).NpcIndex)
        End If
    Next
End Select

End Sub
'Sub CancelarSacrificio(Sacrificado As Integer)
'Dim Sacrificador As Integer

'Sacrificador = UserList(Sacrificado).flags.Sacrificador

'UserList(Sacrificado).flags.Sacrificando = 0
'UserList(Sacrificado).flags.Sacrificador = 0
'UserList(Sacrificador).flags.Sacrificado = 0 '

'Call SendData(ToIndex, Sacrificado, 0, "||¡El sacrificio fue cancelado!" & FONTTYPE_INFO)
'Call SendData(ToIndex, Sacrificador, 0, "||¡El sacrificio fue cancelado!" & FONTTYPE_INFO)

'End Sub
Sub MoveUserChar(userindex As Integer, ByVal nHeading As Byte)
On Error Resume Next
Dim nPos As WorldPos


nPos = UserList(userindex).POS
Call HeadtoPos(nHeading, nPos)
UserList(userindex).Counters.tBoveda = 0
If Not LegalPos(UserList(userindex).POS.Map, nPos.x, nPos.Y, PuedeAtravesarAgua(userindex), UserList(userindex).flags.Privilegios) Or Not CTFPos(userindex, nPos) Then
    Call SendData(ToIndex, userindex, 0, "PU" & UserList(userindex).POS.x & "," & UserList(userindex).POS.Y)
    If MapData(nPos.Map, nPos.x, nPos.Y).userindex Then
        Call EnviaNuevaPosUsuarioPj(userindex, MapData(nPos.Map, nPos.x, nPos.Y).userindex)
    ElseIf MapData(nPos.Map, nPos.x, nPos.Y).NpcIndex Then
        Call EnviaNuevaPosNPC(userindex, MapData(nPos.Map, nPos.x, nPos.Y).NpcIndex)
    End If
    Exit Sub
End If

Call SendData(ToPCAreaButIndexG, userindex, UserList(userindex).POS.Map, ("MP" & UserList(userindex).Char.CharIndex & "," & nPos.x & "," & nPos.Y))
Call EnviaGenteEnNuevoRango(userindex, nHeading)
MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).userindex = 0
UserList(userindex).POS = nPos
UserList(userindex).Char.Heading = nHeading
MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).userindex = userindex
Call DoTileEvents(userindex)

End Sub

Function CTFPos(userindex As Integer, PosCTF As WorldPos)
CTFPos = True
If UserList(userindex).POS.Map <> 196 And UserList(userindex).POS.Map <> 194 Then Exit Function
If UserList(userindex).Faccion.Bando = Real And MapData(PosCTF.Map, PosCTF.x, PosCTF.Y).trigger = 14 Then
        CTFPos = True
        Exit Function
    ElseIf UserList(userindex).Faccion.Bando = Real And MapData(PosCTF.Map, PosCTF.x, PosCTF.Y).trigger = 13 Then
        If TieneObjetos(BANDERAINDEXCRIMI, 1, userindex) Then
            Call Gano(userindex)
        End If
            CTFPos = False
            Exit Function
    ElseIf UserList(userindex).Faccion.Bando = Caos And MapData(PosCTF.Map, PosCTF.x, PosCTF.Y).trigger = 14 Then
        If TieneObjetos(BANDERAINDEXCIUDA, 1, userindex) Then
            Call Gano(userindex)
        End If
            CTFPos = False
            Exit Function
    ElseIf UserList(userindex).Faccion.Bando = Caos And MapData(PosCTF.Map, PosCTF.x, PosCTF.Y).trigger = 13 Then
        CTFPos = True
        Exit Function
End If


End Function

Sub DesequiparItem(userindex As Integer, Slot As Byte)

Call SendData(ToIndex, userindex, 0, "8J" & Slot)

End Sub
Sub EquiparItem(userindex As Integer, Slot As Byte)

Call SendData(ToIndex, userindex, 0, "7J" & Slot)

End Sub

Sub SendUserItem(userindex As Integer, Slot As Byte, JustAmount As Boolean)
Dim MiObj As UserOBJ
Dim Info As String

MiObj = UserList(userindex).Invent.Object(Slot)

If MiObj.OBJIndex Then
    If Not JustAmount Then
        Info = "CSI" & Slot & "," & ObjData(MiObj.OBJIndex).Name & "," & MiObj.Amount & "," & MiObj.Equipped & "," & ObjData(MiObj.OBJIndex).GrhIndex & "," _
        & ObjData(MiObj.OBJIndex).ObjType & "," & Round(ObjData(MiObj.OBJIndex).Valor / 3)
        Select Case ObjData(MiObj.OBJIndex).ObjType
            Case OBJTYPE_WEAPON
                Info = Info & "," & ObjData(MiObj.OBJIndex).MaxHIT & "," & ObjData(MiObj.OBJIndex).MinHIT
            Case OBJTYPE_ARMOUR
                Info = Info & "," & ObjData(MiObj.OBJIndex).SubTipo & "," & ObjData(MiObj.OBJIndex).MaxDef & "," & ObjData(MiObj.OBJIndex).MinDef
            Case OBJTYPE_POCIONES
                Info = Info & "," & ObjData(MiObj.OBJIndex).TipoPocion & "," & ObjData(MiObj.OBJIndex).MaxModificador & "," & ObjData(MiObj.OBJIndex).MinModificador
        End Select
        Call SendData(ToIndex, userindex, 0, Info)
    Else: Call SendData(ToIndex, userindex, 0, "CSO" & Slot & "," & MiObj.Amount)
    End If
Else: Call SendData(ToIndex, userindex, 0, "2H" & Slot)
End If

End Sub
Function NextOpenCharIndex() As Integer
Dim LoopC As Integer

For LoopC = 1 To LastChar + 1
    If CharList(LoopC) = 0 Then
        NextOpenCharIndex = LoopC
        NumChars = NumChars + 1
        If LoopC > LastChar Then LastChar = LoopC
        Exit Function
    End If
Next

End Function
Function NextOpenUser() As Integer

Dim LoopC As Integer
  
For LoopC = 1 To MaxUsers + 1
  If LoopC > MaxUsers Then Exit For
  If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
Next LoopC
  
NextOpenUser = LoopC

End Function

Sub SendUserStatsBox(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "EST" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.MinHP & "," & UserList(userindex).Stats.MaxMAN & "," & UserList(userindex).Stats.MinMAN & "," & UserList(userindex).Stats.MaxSta & "," & UserList(userindex).Stats.MinSta & "," & UserList(userindex).Stats.GLD & "," & UserList(userindex).Stats.ELV & "," & UserList(userindex).Stats.ELU & "," & UserList(userindex).Stats.Exp & "," & UserList(userindex).POS.Map)
End Sub
Sub SendUserHP(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "5A" & UserList(userindex).Stats.MinHP)
End Sub
Sub SendUserMANA(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "5D" & UserList(userindex).Stats.MinMAN)
End Sub
Sub SendUserMAXHP(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "8B" & UserList(userindex).Stats.MaxHP)
End Sub
Sub SendUserMAXMANA(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "9B" & UserList(userindex).Stats.MaxMAN)
End Sub
Sub SendUserSTA(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "5E" & UserList(userindex).Stats.MinSta)
End Sub
Sub SendUserORO(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "5F" & UserList(userindex).Stats.GLD)
End Sub
Sub SendUserEXP(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "5G" & UserList(userindex).Stats.Exp)
End Sub
Sub SendUserMANASTA(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "5H" & UserList(userindex).Stats.MinMAN & "," & UserList(userindex).Stats.MinSta)
End Sub
Sub SendUserHPSTA(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "5I" & UserList(userindex).Stats.MinHP & "," & UserList(userindex).Stats.MinSta)
End Sub
Sub EnviarHambreYsed(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "EHYS" & UserList(userindex).Stats.MaxAGU & "," & UserList(userindex).Stats.MinAGU & "," & UserList(userindex).Stats.MaxHam & "," & UserList(userindex).Stats.MinHam)
End Sub
Sub EnviarHyS(userindex As Integer)
Call SendData(ToIndex, userindex, 0, "5J" & UserList(userindex).Stats.MinAGU & "," & UserList(userindex).Stats.MinHam)
End Sub

Sub SendUserSTAtsTxt(ByVal sendIndex As Integer, userindex As Integer)

Call SendData(ToIndex, sendIndex, 0, "||Estadisticas de: " & UserList(userindex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Nivel: " & UserList(userindex).Stats.ELV & "  EXP: " & UserList(userindex).Stats.Exp & "/" & UserList(userindex).Stats.ELU & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Vitalidad: " & UserList(userindex).Stats.FIT & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "||Salud: " & UserList(userindex).Stats.MinHP & "/" & UserList(userindex).Stats.MaxHP & "  Mana: " & UserList(userindex).Stats.MinMAN & "/" & UserList(userindex).Stats.MaxMAN & "  Vitalidad: " & UserList(userindex).Stats.MinSta & "/" & UserList(userindex).Stats.MaxSta & FONTTYPE_INFO)

If UserList(userindex).Invent.WeaponEqpObjIndex Then
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(userindex).Stats.MinHIT & "/" & UserList(userindex).Stats.MaxHIT & " (" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(userindex).Stats.MinHIT & "/" & UserList(userindex).Stats.MaxHIT & FONTTYPE_INFO)
End If

Call SendData(ToIndex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef + 2 * Buleano(UserList(userindex).Clase = GUERRERO And UserList(userindex).Recompensas(2) = 2) & "/" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef + 2 * Buleano(UserList(userindex).Clase = GUERRERO And UserList(userindex).Recompensas(2) = 2) & FONTTYPE_INFO)

If UserList(userindex).Invent.CascoEqpObjIndex Then
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
Else
    Call SendData(ToIndex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
End If

If UserList(userindex).Invent.EscudoEqpObjIndex Then
    Call SendData(ToIndex, sendIndex, 0, "||(ESCUDO) Defensa extra: " & ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MinDef & " / " & ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MaxDef & FONTTYPE_INFO)
End If

If Len(UserList(userindex).GuildInfo.GuildName) > 0 Then
    Call SendData(ToIndex, sendIndex, 0, "||Clan: " & UserList(userindex).GuildInfo.GuildName & FONTTYPE_INFO)
    If UserList(userindex).GuildInfo.EsGuildLeader = 1 Then
       If UserList(userindex).GuildInfo.ClanFundado = UserList(userindex).GuildInfo.GuildName Then
            Call SendData(ToIndex, sendIndex, 0, "||Status: " & "Fundador/Lider" & FONTTYPE_INFO)
       Else
            Call SendData(ToIndex, sendIndex, 0, "||Status: " & "Lider" & FONTTYPE_INFO)
       End If
    Else
        Call SendData(ToIndex, sendIndex, 0, "||Status: " & UserList(userindex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    Call SendData(ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(userindex).GuildInfo.GuildPoints & FONTTYPE_INFO)
End If

Call SendData(ToIndex, sendIndex, 0, "||Oro: " & UserList(userindex).Stats.GLD & "  Posicion: " & UserList(userindex).POS.x & "," & UserList(userindex).POS.Y & " en mapa " & UserList(userindex).POS.Map & FONTTYPE_INFO)

Call SendData(ToIndex, sendIndex, 0, "||Ciudadanos matados: " & UserList(userindex).Faccion.Matados(Real) & " / Criminales matados: " & UserList(userindex).Faccion.Matados(Caos) & " / Neutrales matados: " & UserList(userindex).Faccion.Matados(Neutral) & FONTTYPE_INFO)

End Sub
Sub SendUserInvTxt(ByVal sendIndex As Integer, userindex As Integer)
On Error Resume Next
Dim j As Byte

Call SendData(ToIndex, sendIndex, 0, "||" & UserList(userindex).Name & FONTTYPE_INFO)
Call SendData(ToIndex, sendIndex, 0, "|| Tiene " & UserList(userindex).Invent.NroItems & " objetos." & FONTTYPE_INFO)

For j = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(j).OBJIndex Then
        Call SendData(ToIndex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(userindex).Invent.Object(j).OBJIndex).Name & " Cantidad:" & UserList(userindex).Invent.Object(j).Amount & FONTTYPE_INFO)
    End If
Next

End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, userindex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(ToIndex, sendIndex, 0, "||" & UserList(userindex).Name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(ToIndex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(userindex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
End Sub
Sub UpdateFuerzaYAg(userindex As Integer)
Dim Fue As Integer
Dim Agi As Integer
Dim Tim As Long

Fue = UserList(userindex).Stats.UserAtributos(fuerza)
If Fue = UserList(userindex).Stats.UserAtributosBackUP(fuerza) Then Fue = 0

Agi = UserList(userindex).Stats.UserAtributos(Agilidad)
If Agi = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) Then Agi = 0


Tim = 45 - TiempoTranscurrido(UserList(userindex).flags.DuracionEfecto)
If Tim < 0 Then Tim = 0
Call SendData(ToIndex, userindex, 0, "EIFYA" & Fue & "," & Agi & "," & Tim & "," & UserList(userindex).Stats.UserAtributosBackUP(fuerza) & "," & UserList(userindex).Stats.UserAtributosBackUP(Agilidad))

End Sub
Sub UpdateUserMap(userindex As Integer)
On Error GoTo ErrorHandler
Dim TempChar As Integer
Dim Map As Integer
Dim x As Integer
Dim Y As Integer
Dim i As Integer

Map = UserList(userindex).POS.Map

Call SendData(ToIndex, userindex, 0, "ET")


For i = 1 To MapInfo(Map).NumUsers
    TempChar = MapInfo(Map).userindex(i)
    Call MakeUserChar(ToIndex, userindex, 0, TempChar, Map, UserList(TempChar).POS.x, UserList(TempChar).POS.Y)
Next


For i = 1 To LastNPC
    If Npclist(i).flags.NPCActive And UserList(userindex).POS.Map = Npclist(i).POS.Map Then
        Call MakeNPCChar(ToIndex, userindex, 0, i, Map, Npclist(i).POS.x, Npclist(i).POS.Y)
    End If
Next


For Y = YMinMapSize To YMaxMapSize
    For x = XMinMapSize To XMaxMapSize
        If MapData(Map, x, Y).OBJInfo.OBJIndex Then
            If ObjData(MapData(Map, x, Y).OBJInfo.OBJIndex).ObjType <> OBJTYPE_ARBOLES Or MapData(Map, x, Y).trigger = 2 Then
                If Y >= 40 Then
                    Y = Y
                End If
                
                Call MakeObj(ToIndex, userindex, 0, MapData(Map, x, Y).OBJInfo, Map, x, Y)
                
                If ObjData(MapData(Map, x, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
                    Call Bloquear(ToIndex, userindex, 0, Map, x, Y, MapData(Map, x, Y).Blocked)
                    Call Bloquear(ToIndex, userindex, 0, Map, x - 1, Y, MapData(Map, x - 1, Y).Blocked)
                End If
            End If
        End If
    Next
Next

Exit Sub
ErrorHandler:
    Call LogError("Error en el sub.UpdateUserMap. Mapa: " & Map & "-" & x & "-" & Y)

End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function
Function EsMascotaCiudadano(NpcIndex As Integer, userindex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser Then
    EsMascotaCiudadano = UserList(userindex).Faccion.Bando = Real
    If EsMascotaCiudadano Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "F0" & UserList(userindex).Name)
End If

End Function
Function EsMascotaCriminal(NpcIndex As Integer, userindex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser Then
    EsMascotaCriminal = Not UserList(userindex).Faccion.Bando = Caos
    If EsMascotaCriminal Then Call SendData(ToIndex, Npclist(NpcIndex).MaestroUser, 0, "F0" & UserList(userindex).Name)
End If

End Function
Sub NpcAtacado(NpcIndex As Integer, userindex As Integer)

Npclist(NpcIndex).flags.AttackedBy = userindex
Call QuitarInvisible(userindex)

If Npclist(NpcIndex).MaestroUser Then Call AllMascotasAtacanUser(userindex, Npclist(NpcIndex).MaestroUser)
If Npclist(NpcIndex).flags.Faccion <> Neutral Then
    If UserList(userindex).Faccion.Ataco(Npclist(NpcIndex).flags.Faccion) = 0 Then UserList(userindex).Faccion.Ataco(Npclist(NpcIndex).flags.Faccion) = 2
End If

Npclist(NpcIndex).Movement = NPCDEFENSA
Npclist(NpcIndex).Hostile = 1

End Sub
Function PuedeApuñalar(userindex As Integer) As Boolean

If UserList(userindex).Invent.WeaponEqpObjIndex Then PuedeApuñalar = ((UserList(userindex).Stats.UserSkills(Apuñalar) >= MIN_APUÑALAR) And (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1)) Or ((UserList(userindex).Clase = ASESINO) And (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1))

End Function
Sub SubirSkill(userindex As Integer, Skill As Integer, Optional Prob As Integer)
On Error GoTo errhandler

If UserList(userindex).flags.Hambre = 1 Or UserList(userindex).flags.Sed = 1 Then Exit Sub

If Prob = 0 Then
    If UserList(userindex).Stats.ELV <= 3 Then
        Prob = 1
    ElseIf UserList(userindex).Stats.ELV > 3 _
        And UserList(userindex).Stats.ELV < 6 Then
        Prob = 1
    ElseIf UserList(userindex).Stats.ELV >= 6 _
        And UserList(userindex).Stats.ELV < 10 Then
        Prob = 2
    ElseIf UserList(userindex).Stats.ELV >= 10 _
        And UserList(userindex).Stats.ELV < 20 Then
        Prob = 2
    Else
        Prob = 2
    End If
End If

If UserList(userindex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

If Int(RandomNumber(1, Prob)) = 2 And UserList(userindex).Stats.UserSkills(Skill) < LevelSkill(UserList(userindex).Stats.ELV).LevelValue Then
    Call AddtoVar(UserList(userindex).Stats.UserSkills(Skill), 1, MAXSKILLPOINTS)
    Call SendData(ToIndex, userindex, 0, "G0" & SkillsNames(Skill) & "," & UserList(userindex).Stats.UserSkills(Skill))
    Call AddtoVar(UserList(userindex).Stats.Exp, 50, MaxExp)
    Call SendData(ToIndex, userindex, 0, "EX" & 50)
    Call SendUserEXP(userindex)
    Call CheckUserLevel(userindex)
End If
Exit Sub

errhandler:
    Call LogError("Error en SubirSkill: " & Err.Description & "-" & UserList(userindex).Name & "-" & SkillsNames(Skill))
End Sub
Sub BajarInvisible(userindex As Integer)

If UserList(userindex).Stats.ELV >= 34 Or UserList(userindex).flags.GolpeoInvi Then
    Call QuitarInvisible(userindex)
Else: UserList(userindex).flags.GolpeoInvi = 1
End If

End Sub
Sub QuitarInvisible(userindex As Integer)

UserList(userindex).Counters.Invisibilidad = 0
UserList(userindex).flags.Invisible = 0
UserList(userindex).flags.GolpeoInvi = 0
UserList(userindex).flags.Oculto = 0
Call SendData(ToMap, 0, UserList(userindex).POS.Map, ("V3" & UserList(userindex).Char.CharIndex & ",0"))

End Sub
Sub UserDie(userindex As Integer)
On Error GoTo ErrorHandler

Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_USERMUERTE)

'If UserList(userindex).flags.Montado = 1 Then Desmontar (userindex)

Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "QDL" & UserList(userindex).Char.CharIndex)

UserList(userindex).Stats.MinHP = 0
UserList(userindex).flags.AtacadoPorNpc = 0
UserList(userindex).flags.AtacadoPorUser = 0
UserList(userindex).flags.Envenenado = 0
UserList(userindex).flags.Muerto = 1

Dim aN As Integer

aN = UserList(userindex).flags.AtacadoPorNpc

If aN Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = 0
End If

If UserList(userindex).flags.Paralizado Then
    Call SendData(ToIndex, userindex, 0, "P8")
    UserList(userindex).flags.Paralizado = 0
End If

If UserList(userindex).flags.Trabajando Then Call SacarModoTrabajo(userindex)

If UserList(userindex).flags.Invisible And UserList(userindex).flags.AdminInvisible = 0 Then
    Call QuitarInvisible(userindex)
End If

If UserList(userindex).flags.Ceguera = 1 Then
  UserList(userindex).Counters.Ceguera = 0
  UserList(userindex).flags.Ceguera = 0
  Call SendData(ToMap, 0, UserList(userindex).POS.Map, "NSEGUE")
End If

If UserList(userindex).flags.Estupidez = 1 Then
  UserList(userindex).Counters.Estupidez = 0
  UserList(userindex).flags.Estupidez = 0
  Call SendData(ToMap, 0, UserList(userindex).POS.Map, "NESTUP")
End If

If UserList(userindex).flags.Descansar Then
    UserList(userindex).flags.Descansar = False
    Call SendData(ToIndex, userindex, 0, "DOK")
End If

If UserList(userindex).flags.Meditando Then
    UserList(userindex).flags.Meditando = False
    Call SendData(ToIndex, userindex, 0, "MEDOK")
End If

Dim ASD As Obj
ASD.Amount = 1


If UserList(userindex).POS.Map = MAP_CTF Or UserList(userindex).POS.Map = MAP_CTC Then
Dim Slotx As Integer
Slotx = Slotx + 1
    If UserList(userindex).Faccion.Bando = Real Then
    
    If TieneObjetos(BANDERAINDEXCRIMI, 1, userindex) = True Then
        Do Until UserList(userindex).Invent.Object(Slotx).OBJIndex = BANDERAINDEXCRIMI
        Slotx = Slotx + 1
        Loop
        Call QuitarVariosItem(userindex, Slotx, 1)
    ASD.OBJIndex = BANDERAINDEXCRIMI
    Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, ASD, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
        
    End If

    Call WarpUserChar(userindex, 206, 12, 20, True)
    ElseIf UserList(userindex).Faccion.Bando = Caos Then
    
        If TieneObjetos(BANDERAINDEXCIUDA, 1, userindex) = True Then
            Do Until UserList(userindex).Invent.Object(Slotx).OBJIndex = BANDERAINDEXCIUDA
            Slotx = Slotx + 1
            Loop
            Call QuitarVariosItem(userindex, Slotx, 1)
        ASD.OBJIndex = BANDERAINDEXCIUDA
        Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, ASD, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
      
        End If

    Call WarpUserChar(userindex, 206, 59, 73, True)
    End If
End If


If UserList(userindex).POS.Map <> 190 And _
UserList(userindex).POS.Map <> 172 And _
UserList(userindex).POS.Map <> 206 And (MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y).trigger <> 6 Or UserList(userindex).POS.Map = 170) Then
    If Not EsNewbie(userindex) Then
        If UserList(userindex).flags.Privilegios > 0 Then Exit Sub
        If UserList(userindex).flags.EnDM = False Then
        Call TirarTodo(userindex)
        End If
    Else: Call TirarTodosLosItemsNoNewbies(userindex)
    End If
End If



If UserList(userindex).Invent.ArmourEqpObjIndex Then Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
If UserList(userindex).Invent.WeaponEqpObjIndex Then Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
If UserList(userindex).Invent.EscudoEqpObjIndex Then Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
If UserList(userindex).Invent.CascoEqpObjIndex Then Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
If UserList(userindex).Invent.HerramientaEqpObjIndex Then Call Desequipar(userindex, UserList(userindex).Invent.HerramientaEqpslot)
If UserList(userindex).Invent.MunicionEqpObjIndex Then Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)

If UserList(userindex).Char.loops = LoopAdEternum Then
    UserList(userindex).Char.FX = 0
    UserList(userindex).Char.loops = 0
End If

If UserList(userindex).flags.Navegando = 0 Then
    UserList(userindex).Char.Body = iCuerpoMuerto
    UserList(userindex).Char.Head = iCabezaMuerto
    UserList(userindex).Char.ShieldAnim = NingunEscudo
    UserList(userindex).Char.WeaponAnim = NingunArma
    UserList(userindex).Char.CascoAnim = NingunCasco
Else
    UserList(userindex).Char.Body = iFragataFantasmal
End If

Dim i As Integer
For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(i) Then
           If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia Then
                Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
           Else
                Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(userindex).MascotasIndex(i)).Movement = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldMovement
                Npclist(UserList(userindex).MascotasIndex(i)).Hostile = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldHostil
                UserList(userindex).MascotasIndex(i) = 0
                UserList(userindex).MascotasType(i) = 0
           End If
    End If
    
Next

If UserList(userindex).POS.Map <> 190 Or UserList(userindex).POS.Map <> 170 Then UserList(userindex).Stats.VecesMurioUsuario = UserList(userindex).Stats.VecesMurioUsuario + 1


UserList(userindex).NroMascotas = 0

Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, val(userindex), UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
If PuedeDestrabarse(userindex) Then Call SendData(ToIndex, userindex, 0, "||Estás encerrado, para destrabarte presiona la tecla Z." & FONTTYPE_INFO)
Call SendUserStatsBox(userindex)

    If MapInfo(UserList(userindex).POS.Map).QuestMod = True Then
        Select Case UserList(userindex).Faccion.Bando
            Case 0
        'Call WarpUserChar(userindex, 1, 50, 50, True)
            Case 1
        Call WarpUserChar(userindex, 34, 50, 50, True)
           Case 2
        Call WarpUserChar(userindex, 1, 50, 50, True)
        End Select
    End If


Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE")

End Sub
Sub ContarMuerte(Muerto As Integer, Atacante As Integer)
If EsNewbie(Muerto) Then Exit Sub

If UserList(Muerto).POS.Map = 190 Or UserList(Muerto).POS.Map = 170 Then Exit Sub

If UserList(Atacante).flags.LastMatado(UserList(Muerto).Faccion.Bando) <> UCase$(UserList(Muerto).Name) Then
    UserList(Atacante).flags.LastMatado(UserList(Muerto).Faccion.Bando) = UCase$(UserList(Muerto).Name)
    Call AddtoVar(UserList(Atacante).Faccion.Matados(UserList(Muerto).Faccion.Bando), 1, 65000)
End If

End Sub

Sub Tilelibre(POS As WorldPos, nPos As WorldPos)


Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
hayobj = False
nPos.Map = POS.Map

Do While Not LegalPos(POS.Map, nPos.x, nPos.Y) Or hayobj
    
    If LoopC > 15 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = POS.Y - LoopC To POS.Y + LoopC
        For tX = POS.x - LoopC To POS.x + LoopC
        
            If LegalPos(nPos.Map, tX, tY) Then
               hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.OBJIndex > 0)
               If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                     nPos.x = tX
                     nPos.Y = tY
                     tX = POS.x + LoopC
                     tY = POS.Y + LoopC
                End If
            End If
        
        Next
    Next
    
    LoopC = LoopC + 1
    
Loop

If Notfound Then
    nPos.x = 0
    nPos.Y = 0
End If

End Sub
Sub AgregarAUsersPorMapa(userindex As Integer)


MapInfo(UserList(userindex).POS.Map).NumUsers = MapInfo(UserList(userindex).POS.Map).NumUsers + 1
If MapInfo(UserList(userindex).POS.Map).NumUsers < 0 Then MapInfo(UserList(userindex).POS.Map).NumUsers = 0

If MapInfo(UserList(userindex).POS.Map).NumUsers = 1 Then
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
    ReDim MapInfo(UserList(userindex).POS.Map).userindex(1 To 1)
Else
    
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
    ReDim Preserve MapInfo(UserList(userindex).POS.Map).userindex(1 To MapInfo(UserList(userindex).POS.Map).NumUsers)
End If


MapInfo(UserList(userindex).POS.Map).userindex(MapInfo(UserList(userindex).POS.Map).NumUsers) = userindex
    
End Sub
Sub QuitarDeUsersPorMapa(userindex As Integer)


MapInfo(UserList(userindex).POS.Map).NumUsers = MapInfo(UserList(userindex).POS.Map).NumUsers - 1
If MapInfo(UserList(userindex).POS.Map).NumUsers < 0 Then MapInfo(UserList(userindex).POS.Map).NumUsers = 0

If MapInfo(UserList(userindex).POS.Map).NumUsers Then
    Dim i As Integer
        
    For i = 1 To MapInfo(UserList(userindex).POS.Map).NumUsers + 1
        
        If MapInfo(UserList(userindex).POS.Map).userindex(i) = userindex Then Exit For
    Next
    
    For i = i To MapInfo(UserList(userindex).POS.Map).NumUsers
        
        MapInfo(UserList(userindex).POS.Map).userindex(i) = MapInfo(UserList(userindex).POS.Map).userindex(i + 1)
    Next
    
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
    ReDim Preserve MapInfo(UserList(userindex).POS.Map).userindex(1 To MapInfo(UserList(userindex).POS.Map).NumUsers)
Else
    ReDim MapInfo(UserList(userindex).POS.Map).userindex(0)
End If
    
End Sub
Sub WarpUserChar(userindex As Integer, Map As Integer, x As Integer, Y As Integer, Optional FX As Boolean = False)

Call SendData(ToMap, 0, UserList(userindex).POS.Map, "QDL" & UserList(userindex).Char.CharIndex)
Call SendData(ToIndex, userindex, UserList(userindex).POS.Map, "QTDL")

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

UserList(userindex).Counters.Protegido = 2
UserList(userindex).flags.Protegido = 3

OldMap = UserList(userindex).POS.Map
OldX = UserList(userindex).POS.x
OldY = UserList(userindex).POS.Y

Call EraseUserChar(ToMap, 0, OldMap, userindex)

UserList(userindex).POS.x = x
UserList(userindex).POS.Y = Y

If OldMap = Map Then
    Call MakeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
    Call SendData(ToIndex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
Else
    Call QuitarDeUsersPorMapa(userindex)
    UserList(userindex).POS.Map = Map
    Call AgregarAUsersPorMapa(userindex)
     
    Call SendData(ToIndex, userindex, 0, "CM" & UserList(userindex).POS.Map & "," & MapInfo(UserList(userindex).POS.Map).MapVersion & "," & MapInfo(UserList(userindex).POS.Map).Name & "," & MapInfo(UserList(userindex).POS.Map).TopPunto & "," & MapInfo(UserList(userindex).POS.Map).LeftPunto)
    If MapInfo(Map).Music <> MapInfo(OldMap).Music Then Call SendData(ToIndex, userindex, 0, "TM" & MapInfo(Map).Music)

    Call MakeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
    Call SendData(ToIndex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
End If

Call UpdateUserMap(userindex)

If FX And UserList(userindex).flags.AdminInvisible = 0 And Not UserList(userindex).flags.Meditando Then
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_WARP)
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXWARP & "," & 0)
End If
Dim i As Integer
For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(i) Then
           If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia Then
                Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
           Else
                Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = 0
                Npclist(UserList(userindex).MascotasIndex(i)).Movement = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldMovement
                Npclist(UserList(userindex).MascotasIndex(i)).Hostile = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldHostil
                UserList(userindex).MascotasIndex(i) = 0
                UserList(userindex).MascotasType(i) = 0
           End If
    End If
    
Next
UserList(userindex).NroMascotas = 0
'WarpMascotas (userindex)

End Sub
Sub WarpMascotas(userindex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
Dim PetTypes(1 To MAXMASCOTAS) As Integer
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer

NroPets = UserList(userindex).NroMascotas

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(i) Then
        PetRespawn(i) = Npclist(UserList(userindex).MascotasIndex(i)).flags.Respawn = 0
        If PetRespawn(i) Then
            PetTypes(i) = UserList(userindex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
        Else
            PetTypes(i) = UserList(userindex).MascotasType(i)
            PetTiempoDeVida(i) = 0
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
        End If
    End If
Next

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If PetTypes(i) Then
        UserList(userindex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(userindex).POS, False, PetRespawn(i))
        UserList(userindex).MascotasType(i) = PetTypes(i)
        
        If UserList(userindex).MascotasIndex(i) = MAXNPCS Then
                UserList(userindex).MascotasIndex(i) = 0
                UserList(userindex).MascotasType(i) = 0
                If UserList(userindex).NroMascotas Then UserList(userindex).NroMascotas = UserList(userindex).NroMascotas - 1
                Exit Sub
        End If
        Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = userindex
        Npclist(UserList(userindex).MascotasIndex(i)).Movement = SIGUE_AMO
        Npclist(UserList(userindex).MascotasIndex(i)).Target = 0
        Npclist(UserList(userindex).MascotasIndex(i)).TargetNpc = 0
        Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
        Call QuitarNPCDeLista(Npclist(UserList(userindex).MascotasIndex(i)).Numero, UserList(userindex).POS.Map)
        Call FollowAmo(UserList(userindex).MascotasIndex(i))
    End If
Next

UserList(userindex).NroMascotas = NroPets

End Sub
Sub Cerrar_Usuario(userindex As Integer)

If UserList(userindex).flags.UserLogged And Not UserList(userindex).Counters.Saliendo Then
    UserList(userindex).Counters.Saliendo = True
    UserList(userindex).Counters.Salir = Timer - 8 * Buleano(UserList(userindex).Clase = PIRATA And UserList(userindex).Recompensas(3) = 2)
    Call SendData(ToIndex, userindex, 0, "SAL1")
    Call SendData(ToIndex, userindex, 0, "1Z" & IntervaloCerrarConexion - 8 * Buleano(UserList(userindex).Clase = PIRATA And UserList(userindex).Recompensas(3) = 2))
End If
    
End Sub
