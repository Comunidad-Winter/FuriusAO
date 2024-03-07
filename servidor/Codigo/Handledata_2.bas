Attribute VB_Name = "Handledata_2"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Public Sub HandleData2(userindex As Integer, rdata As String, Procesado As Boolean)
Dim LoopC As Integer, tIndex As Integer, N As Integer, x As Integer, Y As Integer, tInt As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tLong As Long

Procesado = True




Select Case Left$(UCase$(rdata), 2)

Case "#¬"
     Call EnviarRecom(userindex)
     Exit Sub

Case "#X"
If Abierto = False Then Exit Sub
If UserList(userindex).flags.YaVoto Then SendData ToIndex, userindex, 0, "||Ya has votado" & FONTTYPE_CELESTE: Exit Sub
UserList(userindex).flags.YaVoto = True
SIs = SIs + 1
Call SendData(ToIndex, userindex, 0, "||Gracias por votar" & FONTTYPE_CELESTE)
Exit Sub


Case "#Z"
If Abierto = False Then Exit Sub
If UserList(userindex).flags.YaVoto Then SendData ToIndex, userindex, 0, "||Ya has votado" & FONTTYPE_CELESTE: Exit Sub
UserList(userindex).flags.YaVoto = True
NOs = NOs + 1
Call SendData(ToIndex, userindex, 0, "||Gracias por votar" & FONTTYPE_CELESTE)
Exit Sub




    Case "#*"
        rdata = Right$(rdata, Len(rdata) - 2)
        tIndex = NameIndex(rdata)
        If tIndex Then
            If UserList(tIndex).flags.Privilegios < 2 Then
                Call SendData(ToIndex, userindex, 0, "||El jugador " & UserList(tIndex).Name & " se encuentra online." & FONTTYPE_INFO)
            Else: Call SendData(ToIndex, userindex, 0, "1A")
            End If
        Else: Call SendData(ToIndex, userindex, 0, "1A")
        End If
        Exit Sub
    Case "#]"
        rdata = Right$(rdata, Len(rdata) - 2)
        Call TirarRuleta(userindex, rdata)
    
        Exit Sub
    Case "#}"
        UserList(userindex).flags.MesaCasino = 0
        Call SendUserORO(userindex)
        Exit Sub
        
    Case "^A"
        rdata = Right$(rdata, Len(rdata) - 2)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & ": " & rdata & FONTTYPE_FIGHT)
        Exit Sub
    
    Case "#$"
        rdata = Right$(rdata, Len(rdata) - 2)
        If UserList(userindex).flags.Privilegios < 2 Then Exit Sub
        x = ReadField(1, rdata, 44)
        Y = ReadField(2, rdata, 44)
        N = MapaPorUbicacion(x, Y)
        If N Then Call WarpUserChar(userindex, N, 50, 50, True)
        Call LogGM(UserList(userindex).Name, "Se transporto por mapa a Mapa" & mapa & " X:" & x & " Y:" & Y, (UserList(userindex).flags.Privilegios = 1))
        Exit Sub
    
    Case "#A"
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        If Not UserList(userindex).flags.Meditando And UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN Then Exit Sub
        Call SendData(ToIndex, userindex, 0, "MEDOK")
        If Not UserList(userindex).flags.Meditando Then
           Call SendData(ToIndex, userindex, 0, "7M")
        Else
           Call SendData(ToIndex, userindex, 0, "D9")
        End If
        UserList(userindex).flags.Meditando = Not UserList(userindex).flags.Meditando
        
        If UserList(userindex).flags.Meditando Then
            UserList(userindex).Counters.tInicioMeditar = Timer
            Call SendData(ToIndex, userindex, 0, "8M" & TIEMPO_INICIOMEDITAR)

'furiusao nueva meditacion nivel 43 - 45
                    UserList(userindex).Char.loops = LoopAdEternum
                    If UserList(userindex).Stats.ELV < 15 Then
                        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
                        UserList(userindex).Char.FX = FXMEDITARCHICO
                    ElseIf UserList(userindex).Stats.ELV < 30 Then
                        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
                        UserList(userindex).Char.FX = FXMEDITARMEDIANO
                    ElseIf UserList(userindex).Stats.ELV < 43 Then
                            Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
                            UserList(userindex).Char.FX = FXMEDITARGRANDE
                    Else
                            Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXMEDITAREXTRAGRANDE & "," & LoopAdEternum)
                            UserList(userindex).Char.FX = FXMEDITAREXTRAGRANDE
                    End If
            
            Else
                UserList(userindex).Counters.bPuedeMeditar = False
                
                UserList(userindex).Char.FX = 0
                UserList(userindex).Char.loops = 0
                Call SendData(ToMap, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
        End If
        Exit Sub
    'furiusao nueva meditacion nivel 43 - 45
    Case "#B"
        If UserList(userindex).flags.Paralizado Then Exit Sub
        If UserList(userindex).flags.EnTorneo Then Exit Sub
        If UserList(userindex).flags.EnReto Then Exit Sub
        If (Not MapInfo(UserList(userindex).POS.Map).Pk And TiempoTranscurrido(UserList(userindex).Counters.LastRobo) > 10) Or UserList(userindex).flags.Privilegios > 1 Then
            Call SendData(ToIndex, userindex, 0, "FINOK")
            Call CloseSocket(userindex)
            Exit Sub
        End If
        
        Call Cerrar_Usuario(userindex)
        
        Exit Sub

    Case "#C"
        If CanCreateGuild(userindex) Then Call SendData(ToIndex, userindex, 0, "SHOWFUN" & UserList(userindex).Faccion.Bando)
        Exit Sub
    
    Case "#D"
        Call SendData(ToIndex, userindex, 0, "7L")
        Exit Sub
    
    Case "#E"
        Call SendData(ToIndex, userindex, 0, "7L")
        Exit Sub
    
    Case "#F"
        Call SendData(ToIndex, userindex, 0, "7L")
        Exit Sub
        

    Case "#G"
        
        If UserList(userindex).flags.Muerto Then
                  Call SendData(ToIndex, userindex, 0, "MU")
                  Exit Sub
        End If
        
        If UserList(userindex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, userindex, 0, "ZP")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 3 Then
                  Call SendData(ToIndex, userindex, 0, "DL")
                  Exit Sub
        End If
        If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
        Or UserList(userindex).flags.Muerto Then Exit Sub

        Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex)
        Exit Sub
    Case "#H"
         
         If UserList(userindex).flags.Muerto Then
                      Call SendData(ToIndex, userindex, 0, "MU")
                      Exit Sub
         End If
         
         If UserList(userindex).flags.TargetNpc = 0 Then
                  Call SendData(ToIndex, userindex, 0, "ZP")
                  Exit Sub
         End If
         If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 10 Then
                      Call SendData(ToIndex, userindex, 0, "DL")
                      Exit Sub
         End If
         If Npclist(UserList(userindex).flags.TargetNpc).MaestroUser <> _
            userindex Then Exit Sub
         Npclist(UserList(userindex).flags.TargetNpc).Movement = ESTATICO
         Call Expresar(UserList(userindex).flags.TargetNpc, userindex)
         Exit Sub
    Case "#I"
        
        If UserList(userindex).flags.Muerto Then
                  Call SendData(ToIndex, userindex, 0, "MU")
                  Exit Sub
        End If
        
        If UserList(userindex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, userindex, 0, "ZP")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 10 Then
                  Call SendData(ToIndex, userindex, 0, "DL")
                  Exit Sub
        End If
        If Npclist(UserList(userindex).flags.TargetNpc).MaestroUser <> _
          userindex Then Exit Sub
        Call FollowAmo(UserList(userindex).flags.TargetNpc)
        Call Expresar(UserList(userindex).flags.TargetNpc, userindex)
        Exit Sub
    Case "#J"
        
        If UserList(userindex).flags.Muerto Then
                  Call SendData(ToIndex, userindex, 0, "MU")
                  Exit Sub
        End If
        
        If UserList(userindex).flags.TargetNpc = 0 Then
              Call SendData(ToIndex, userindex, 0, "ZP")
              Exit Sub
        End If
        If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 10 Then
                  Call SendData(ToIndex, userindex, 0, "DL")
                  Exit Sub
        End If
        If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
        Call EnviarListaCriaturas(userindex, UserList(userindex).flags.TargetNpc)
        Exit Sub
    Case "#K"
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        If HayOBJarea(UserList(userindex).POS, FOGATA) Then
                Call SendData(ToIndex, userindex, 0, "DOK")
                If Not UserList(userindex).flags.Descansar Then
                    Call SendData(ToIndex, userindex, 0, "3M")
                Else
                    Call SendData(ToIndex, userindex, 0, "4M")
                End If
                UserList(userindex).flags.Descansar = Not UserList(userindex).flags.Descansar
        Else
                If UserList(userindex).flags.Descansar Then
                    Call SendData(ToIndex, userindex, 0, "4M")
                    
                    UserList(userindex).flags.Descansar = False
                    Call SendData(ToIndex, userindex, 0, "DOK")
                    Exit Sub
                End If
                Call SendData(ToIndex, userindex, 0, "6M")
        End If
        Exit Sub

    Case "#L"
       
       If UserList(userindex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, userindex, 0, "ZP")
           Exit Sub
       End If
       
       If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_REVIVIR _
       Or UserList(userindex).flags.Muerto <> 1 Then Exit Sub
       If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 10 Then
           Call SendData(ToIndex, userindex, 0, "DL")
           Exit Sub
       End If

       Call RevivirUsuarioNPC(userindex)
       Call SendData(ToIndex, userindex, 0, "RZ")
       Exit Sub
    Case "#M"
       
       If UserList(userindex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, userindex, 0, "ZP")
           Exit Sub
       End If
       If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_REVIVIR _
       Or UserList(userindex).flags.Muerto Then Exit Sub
       If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 10 Then
           Call SendData(ToIndex, userindex, 0, "DL")
           Exit Sub
       End If
       UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
       Call SendUserHP(userindex)
       Exit Sub
    Case "#N"
        If UserList(userindex).flags.Muerto Then Exit Sub
        Call EnviarSubclase(userindex)
        Exit Sub
    Case "#O"
        If PuedeRecompensa(userindex) And Not UserList(userindex).flags.Muerto Then _
        Call SendData(ToIndex, userindex, 0, "RELON" & UserList(userindex).Clase & "," & PuedeRecompensa(userindex))
    Exit Sub
    Case "#P"
        If UserList(userindex).flags.Privilegios > 0 Then
            For LoopC = 1 To LastUser
                If Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.Privilegios <= 1 Then
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            Next
            If Len(tStr) > 0 Then
                tStr = Left$(tStr, Len(tStr) - 2)
                tStr = "ONLINE:" & tStr
                Call SendData(ToIndex, userindex, 0, "||" & tStr & FONTTYPE_INFO)
                Call SendData(ToIndex, userindex, 0, "4L" & NumNoGMs)
            Else
                Call SendData(ToIndex, userindex, 0, "6L")
            End If
        Else
           Call SendData(ToIndex, userindex, 0, "||Este comando ya no está disponible. La cantidad de users online está abajo de la pantalla." & FONTTYPE_INFO)
        End If
        Exit Sub
    Case "#Q"
        Call SendUserSTAtsTxt(userindex, userindex)
        Exit Sub
    Case "#R"
        If UserList(userindex).Counters.Pena Then
            Call SendData(ToIndex, userindex, 0, "9M" & CalcularTiempoCarcel(userindex))
        Else
            Call SendData(ToIndex, userindex, 0, "2N")
        End If
        Exit Sub
    Case "#S"
        If UserList(userindex).flags.TargetUser Then
        If UserList(userindex).flags.TargetUser = userindex Then Exit Sub
            If MapData(UserList(UserList(userindex).flags.TargetUser).POS.Map, UserList(UserList(userindex).flags.TargetUser).POS.x, UserList(UserList(userindex).flags.TargetUser).POS.Y).OBJInfo.OBJIndex > 0 And _
            UserList(UserList(userindex).flags.TargetUser).flags.Muerto Then
                Call SendData(ToAdmins, 0, 0, "8T" & UserList(userindex).Name & "," & UserList(UserList(userindex).flags.TargetUser).Name)
                'Call SendData(ToIndex, UserList(userindex).flags.TargetUser, 0, "!!Fuiste echado por mantenerte sobre un item estando muerto.")
                Call SendData(ToIndex, UserList(userindex).flags.TargetUser, 0, "FINOK")
                Call CloseSocket(UserList(userindex).flags.TargetUser)
            End If
        End If
        Exit Sub
    Case "#>"
    If UserList(userindex).flags.EnTorneo Then Exit Sub
    If UserList(userindex).flags.EnReto Then Exit Sub
    InscribirUsuario (userindex)


    Case "#T"
        If EnTorneo Then
        
            If Torneo.ClaseUnica <> "" And Torneo.ClaseUnica <> "TODAS" Then
            If UCase$(ListaClases((UserList(userindex).Clase))) <> UCase$(Torneo.ClaseUnica) Then
            Call SendData(ToIndex, userindex, 0, "||Tu clase no puede participar de este torneo" & FONTTYPE_BLANCO)
            Exit Sub
            End If
            End If

        
            If Torneo.NivelMinimo > UserList(userindex).Stats.ELV Then
            Call SendData(ToIndex, userindex, 0, "||Tu nivel no te permite participar de este torneo" & FONTTYPE_BLANCO)
            Exit Sub
            End If
        
            If UserList(userindex).Stats.GLD < Torneo.Precio Then
            Recaudado = Recaudado + Torneo.Precio
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - Torneo.Precio
            Call SendUserORO(userindex)
            End If
      
        
            Dim jugadores As Integer
            jugadores = val(GetVar(App.Path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD"))
            Dim jugador As Integer
            For jugador = 1 To jugadores
                If UCase$(GetVar(App.Path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugador)) = UCase$(UserList(userindex).Name) & ":" & UserList(userindex).Stats.ELV Then Exit Sub
            Next
            Call WriteVar(App.Path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD", jugadores + 1)
            Call WriteVar(App.Path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugadores + 1, UserList(userindex).Name & ":" & UserList(userindex).Stats.ELV)
            Call SendData(ToIndex, userindex, 0, "9T")
            Call SendData(ToAdmins, 0, 0, "2U" & UserList(userindex).Name)
        End If
        Exit Sub

    Case "#U"
        Dim NpcIndex As Integer
        Dim theading As Byte
        Dim atra1 As Integer
        Dim atra2 As Integer
        Dim atra3 As Integer
        Dim atra4 As Integer

        If Not LegalPos(UserList(userindex).POS.Map, UserList(userindex).POS.x - 1, UserList(userindex).POS.Y) And _
        Not LegalPos(UserList(userindex).POS.Map, UserList(userindex).POS.x + 1, UserList(userindex).POS.Y) And _
        Not LegalPos(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1) And _
        Not LegalPos(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y + 1) Then
            If UserList(userindex).flags.Muerto Then
                If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x - 1, UserList(userindex).POS.Y).NpcIndex Then
                    atra1 = MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x - 1, UserList(userindex).POS.Y).NpcIndex
                    theading = WEST
                    Call MoveNPCChar(atra1, theading)
                End If
                If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x + 1, UserList(userindex).POS.Y).NpcIndex Then
                    atra2 = MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x + 1, UserList(userindex).POS.Y).NpcIndex
                    theading = EAST
                    Call MoveNPCChar(atra2, theading)
                End If
                If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).NpcIndex Then
                    atra3 = MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y - 1).NpcIndex
                    theading = NORTH
                    Call MoveNPCChar(atra3, theading)
                End If
                If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y + 1).NpcIndex Then
                    atra4 = MapData(UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y + 1).NpcIndex
                    theading = SOUTH
                    Call MoveNPCChar(atra4, theading)
                 End If
            End If
        End If
        Exit Sub
        
    Case "#V"
        
        If UserList(userindex).flags.Muerto Then
                  Call SendData(ToIndex, userindex, 0, "MU")
                  Exit Sub
        End If
        
        If UserList(userindex).flags.EnDM = True Then Exit Sub
                
                
        If UserList(userindex).flags.Privilegios = 1 Then
            Exit Sub
        End If
        
        If UserList(userindex).flags.TargetNpc Then
              
              If Npclist(UserList(userindex).flags.TargetNpc).Comercia = 0 Then
                 If Len(Npclist(UserList(userindex).flags.TargetNpc).Desc) > 0 Then Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "3Q" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
                 Exit Sub
              End If
              If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 3 Then
                  Call SendData(ToIndex, userindex, 0, "DL")
                  Exit Sub
              End If
              
              Call IniciarComercioNPC(userindex)
         

        ElseIf UserList(userindex).flags.TargetUser Then
            
            
            If UserList(UserList(userindex).flags.TargetUser).flags.Muerto Then
                Call SendData(ToIndex, userindex, 0, "4U")
                Exit Sub
            End If
            
            If UserList(userindex).flags.TargetUser = userindex Then
                Call SendData(ToIndex, userindex, 0, "5U")
                Exit Sub
            End If
            
            If UserList(userindex).POS.Map = 66 Then
                Call SendData(ToIndex, userindex, 0, "||No puedes comerciar en la cárcel." & FONTTYPE_INFO)
                Exit Sub
            End If


            If Distancia(UserList(UserList(userindex).flags.TargetUser).POS, UserList(userindex).POS) > 3 Then
                Call SendData(ToIndex, userindex, 0, "DL")
                Exit Sub
            End If
            
            If UserList(UserList(userindex).flags.TargetUser).flags.Comerciando And _
                UserList(UserList(userindex).flags.TargetUser).ComUsu.DestUsu <> userindex Then
                Call SendData(ToIndex, userindex, 0, "6U")
                Exit Sub
            End If
            
            UserList(userindex).ComUsu.DestUsu = UserList(userindex).flags.TargetUser
            UserList(userindex).ComUsu.DestNick = UserList(UserList(userindex).flags.TargetUser).Name
            UserList(userindex).ComUsu.Cant = 0
            UserList(userindex).ComUsu.Objeto = 0
            UserList(userindex).ComUsu.Acepto = False
            
            
            Call IniciarComercioConUsuario(userindex, UserList(userindex).flags.TargetUser)

        Else
            Call SendData(ToIndex, userindex, 0, "ZP")
        End If
        Exit Sub
    
    
    Case "#W"
        
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        If UserList(userindex).flags.Bovediando = 0 Then
            If UserList(userindex).flags.TargetNpc = 0 Then
                Call SendData(ToIndex, userindex, 0, "ZP")
                Exit Sub
            End If
            
            If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 10 Then
                Call SendData(ToIndex, userindex, 0, "DL")
                Exit Sub
            End If
            
            If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
        End If
        Call IniciarDeposito(userindex)
        'Call SendData(ToIndex, userindex, 0, "SHWBP")
        Exit Sub

    Case "#Y"
    
    
        If UserList(userindex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, userindex, 0, "ZP")
            Exit Sub
        End If
        
        If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(userindex).flags.Muerto Then Exit Sub
       
        If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 4 Then
            Call SendData(ToIndex, userindex, 0, "DL")
            Exit Sub
        End If
       
        If ClaseBase(UserList(userindex).Clase) Or ClaseTrabajadora(UserList(userindex).Clase) Then Exit Sub
       
        Call Enlistar(userindex, Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion)
       
        Exit Sub

    Case "#1"
        
        If UserList(userindex).flags.TargetNpc = 0 Then
            Call SendData(ToIndex, userindex, 0, "ZP")
            Exit Sub
        End If
        If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(userindex).flags.Muerto Or Not Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion Then Exit Sub
        If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 4 Then
            Call SendData(ToIndex, userindex, 0, "DL")
            Exit Sub
        End If

        If UserList(userindex).Faccion.Bando <> Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion Then
            Call SendData(ToIndex, userindex, 0, Mensajes(Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion, 16) & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
        Call Recompensado(userindex)
        Exit Sub
        
    Case "#5"
        rdata = Right$(rdata, Len(rdata) - 3)
        
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "M4")
            Exit Sub
        End If
        
        If Not AsciiValidos(rdata) Then
            Call SendData(ToIndex, userindex, 0, "7U")
            Exit Sub
        End If
        
        If Len(rdata) > 80 Then
            Call SendData(ToIndex, userindex, 0, "||La descripción debe tener menos de 80 cáracteres de largo." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(userindex).flags.Silenciado > 0 Then Exit Sub
        UserList(userindex).Desc = rdata
        Call SendData(ToIndex, userindex, 0, "8U")
        Exit Sub
        
    Case "#6 "
        rdata = Right$(rdata, Len(rdata) - 3)
        Call ComputeVote(userindex, rdata)
        Exit Sub
            
    Case "#7"
        Call SendData(ToIndex, userindex, 0, "||Este comando ya no anda, para hablar por tu clan presiona la tecla 3 y habla normalmente." & FONTTYPE_INFO)
        Exit Sub

    Case "#8"
        Call SendData(ToIndex, userindex, 0, "||Este comando ya no se usa, pon /PASSWORD para cambiar tu password." & FONTTYPE_INFO)
        Exit Sub
        
    Case "#!"
        If PuedeFaccion(userindex) Then Call SendData(ToIndex, userindex, 0, "4&")
        Exit Sub
        
    Case "#9"
        rdata = Right$(rdata, Len(rdata) - 3)
        tLong = CLng(val(rdata))
        If tLong > 32000 Then tLong = 32000
        N = tLong
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
        ElseIf UserList(userindex).flags.TargetNpc = 0 Then
            
            Call SendData(ToIndex, userindex, 0, "ZP")
        ElseIf Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 10 Then
            Call SendData(ToIndex, userindex, 0, "DL")
        ElseIf Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_APOSTADOR Then
            Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
        ElseIf N < 1 Then
            Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
        ElseIf N > 5000 Then
            Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
        ElseIf UserList(userindex).Stats.GLD < N Then
            Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
        Else
            If RandomNumber(1, 100) <= 47 Then
                UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + N
                Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
                
                Apuestas.Ganancias = Apuestas.Ganancias + N
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            Else
                UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - N
                Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
            
                Apuestas.Perdidas = Apuestas.Perdidas + N
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            End If
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call SendUserORO(userindex)
        End If
        Exit Sub
        
    Case "#/"
        rdata = Right$(rdata, Len(rdata) - 3)
        tIndex = NameIndex(ReadField(1, rdata, 32))
        If tIndex = 0 Then Exit Sub
        If UserList(tIndex).flags.Privilegios > 0 Then Exit Sub
        'Call SendData(ToIndex, userindex, 0, "||No puedes ignorar GMs!" & FONTTYPE_INFO)
        'Exit Sub
        'End If
        If ReadField(2, rdata, 32) = "0" Then
            Call SendData(ToIndex, tIndex, 0, "||" & UserList(userindex).Name & " te ha dejado de ignorar." & FONTTYPE_INFO)
        Else: Call SendData(ToIndex, tIndex, 0, "||" & UserList(userindex).Name & " te empezó a ignorar." & FONTTYPE_INFO)
        End If
        Exit Sub
        
    Case "#0"
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
         

        
         If UserList(userindex).flags.Bovediando = 0 Then
                 If UserList(userindex).flags.TargetNpc = 0 Then
                    Call SendData(ToIndex, userindex, 0, "ZP")
                    Exit Sub
                End If
                If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
                 
                If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 10 Then
                    Call SendData(ToIndex, userindex, 0, "DL")
                    Exit Sub
                End If
         End If
         
        rdata = Right$(rdata, Len(rdata) - 3)
        
        If val(rdata) > 0 Then
            If val(rdata) > UserList(userindex).Stats.Banco Then rdata = UserList(userindex).Stats.Banco
            UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco - val(rdata)
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + val(rdata)
        If UserList(userindex).flags.Bovediando = 0 Then
            Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, userindex, 0, "||Tenés " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & FONTTYPE_INFO)
        End If
        End If
         
        Call SendUserORO(userindex)
         
        Exit Sub
    Case "#;"
        UserList(userindex).flags.Bovediando = 0
        
    Case "#Ñ"
        
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If

            If UserList(userindex).flags.Bovediando = 0 Then
            
                If UserList(userindex).flags.TargetNpc = 0 Then
                    Call SendData(ToIndex, userindex, 0, "ZP")
                    Exit Sub
                End If
                
                If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 10 Then
                    Call SendData(ToIndex, userindex, 0, "DL")
                    Exit Sub
                End If
                
                If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Or UserList(userindex).flags.Muerto Then Exit Sub
                
                If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 10 Then
                      Call SendData(ToIndex, userindex, 0, "DL")
                      Exit Sub
                End If
            End If
        rdata = Right$(rdata, Len(rdata) - 3)
        
        If CLng(val(rdata)) > 0 Then
            If CLng(val(rdata)) > UserList(userindex).Stats.GLD Then rdata = UserList(userindex).Stats.GLD
            UserList(userindex).Stats.Banco = UserList(userindex).Stats.Banco + val(rdata)
            UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(rdata)
            If UserList(userindex).flags.Bovediando = 0 Then
            Call SendData(ToIndex, userindex, 0, "3Q" & vbWhite & "°" & "Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
            Else
            Call SendData(ToIndex, userindex, 0, "||Tenes " & UserList(userindex).Stats.Banco & " monedas de oro en tu cuenta." & FONTTYPE_INFO)
            End If
        End If
    
        Call SendUserORO(userindex)
        
        Exit Sub
        
    Case "#2"
        If Len(UserList(userindex).GuildInfo.GuildName) > 0 Then
            If UserList(userindex).GuildInfo.EsGuildLeader And UserList(userindex).flags.Privilegios < 2 Then
                Call SendData(ToIndex, userindex, 0, "4V")
                Exit Sub
            End If
        Else
            Call SendData(ToIndex, userindex, 0, "5V")
            Exit Sub
        End If
        
        Call SendData(ToGuildMembers, userindex, 0, "6V" & UserList(userindex).Name)
        Call SendData(ToIndex, userindex, 0, "7V")
        
        Dim oGuild As cGuild
        
        Set oGuild = FetchGuild(UserList(userindex).GuildInfo.GuildName)
        
        If oGuild Is Nothing Then Exit Sub
        
        For i = 1 To LastUser
            If UserList(i).GuildInfo.GuildName = oGuild.GuildName Then UserList(i).flags.InfoClanEstatica = 0
        Next
        
        UserList(userindex).GuildInfo.GuildPoints = 0
        UserList(userindex).GuildInfo.GuildName = ""
        Call oGuild.RemoveMember(UserList(userindex).Name)
        
        Call UpdateUserChar(userindex)
        
        Exit Sub
      
    Case "#4"

        If UserList(userindex).flags.TargetNpc = 0 Then
           Call SendData(ToIndex, userindex, 0, "ZP")
           Exit Sub
       End If
       
        If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_NOBLE Or UserList(userindex).flags.Muerto Or Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion = 0 Then Exit Sub
        
        If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 4 Then
            Call SendData(ToIndex, userindex, 0, "DL")
            Exit Sub
        End If
        
        If UserList(userindex).Faccion.Bando <> Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion Then Exit Sub
        
        If Len(UserList(userindex).GuildInfo.GuildName) > 0 Then
            Call SendData(ToIndex, userindex, 0, Mensajes(UserList(userindex).Faccion.Bando, 23) & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
        
        Call SendData(ToIndex, userindex, 0, Mensajes(Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion, 18) & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))

        UserList(userindex).Faccion.Bando = Neutral
        UserList(userindex).Faccion.Jerarquia = 0
        Call UpdateUserChar(userindex)
Exit Sub

Case "#3"
    If Len(UserList(userindex).GuildInfo.GuildName) = 0 Then
        Call SendData(ToIndex, userindex, 0, "5V")
        Exit Sub
    End If
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).GuildInfo.GuildName = UserList(userindex).GuildInfo.GuildName Then
            tStr = tStr & UserList(LoopC).Name & ", "
        End If
    Next
    
    If Len(tStr) > 0 Then
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, userindex, 0, "||Miembros de tu clan online:" & tStr & "." & FONTTYPE_GUILD)
    Else: Call SendData(ToIndex, userindex, 0, "8V")
    End If
    Exit Sub
    


    End Select



    Procesado = False
End Sub
