Attribute VB_Name = "Acciones"
        
Public Cruz As Integer
Public Gema As Integer
Option Explicit
Sub ExtraObjs()

Cruz = UBound(ObjData) - 1
ObjData(Cruz).Name = "Cruz del Sacrificio"
ObjData(Cruz).GrhIndex = 116

Gema = UBound(ObjData)
ObjData(Gema).Name = "Piedra filosofal incompleta"
ObjData(Gema).GrhIndex = 705
    
End Sub
Sub Accion(userindex As Integer, Map As Integer, X As Integer, Y As Integer)
On Error Resume Next

If Not InMapBounds(X, Y) Then Exit Sub
   
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
'ACA LE DAMOS EN LA CABEZAAAAAAA QUIZAS?
Dim SegundaPos As Integer
NuevaPos:

If MapData(Map, X, Y).NpcIndex Then

        If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 10 Then
            Call SendData(ToIndex, userindex, 0, "DL")
            Exit Sub
        End If
        
    If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_REVIVIR Then
        If UserList(userindex).flags.Muerto Then
            Call RevivirUsuarioNPC(userindex)
            Call SendData(ToIndex, userindex, 0, "RZ")
        Else
            UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
            Call SendUserHP(userindex)
        End If
        Exit Sub
        
    End If
    
    
    
    If UserList(userindex).flags.Muerto Then
        Call SendData(ToIndex, userindex, 0, "MU")
        Exit Sub
    End If


    If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_CARCEL Then
        If UserList(userindex).Counters.PenaMinar = 0 Then Exit Sub
        If TieneObjetos(HIERRO_MINA, val(UserList(userindex).Counters.PenaMinar), userindex) Then
        UserList(userindex).Counters.PenaMinar = 0
        UserList(userindex).Counters.TiempoPena = 0
        UserList(userindex).flags.Encarcelado = 0
        UserList(userindex).Counters.Pena = 0
        If UserList(userindex).flags.Trabajando Then Call SacarModoTrabajo(userindex)
        If UserList(userindex).POS.Map = Prision.Map Then
            Call WarpUserChar(userindex, Libertad.Map, Libertad.X, Libertad.Y, True)
            Call SendData(ToIndex, userindex, 0, "4P")
        End If
        Call SaleCarcelPique(userindex)
        Else
        Call SendData(ToIndex, userindex, 0, "||Carcel> Necesitas minar " & UserList(userindex).Counters.PenaMinar & " unidades de hierro." & FONTTYPE_BLANCO)
        End If
        Exit Sub
    End If

    If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_BANQUERO Then
       ' Call IniciarDeposito(userindex)
        Call SendData(ToIndex, userindex, 0, "SHWBP")
        Exit Sub
    End If
    
    If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_TIENDA Then
        If Npclist(MapData(Map, X, Y).NpcIndex).flags.TiendaUser > 0 And Npclist(MapData(Map, X, Y).NpcIndex).flags.TiendaUser <> userindex Then
            Call IniciarComercioTienda(userindex, MapData(Map, X, Y).NpcIndex)
        Else
            Call IniciarAlquiler(userindex)
        End If
        Exit Sub
    End If
    
    If Npclist(MapData(Map, X, Y).NpcIndex).Comercia Then
        Call IniciarComercioNPC(userindex)
        Exit Sub
    End If
    
    If Npclist(MapData(Map, X, Y).NpcIndex).flags.Apostador Then
        UserList(userindex).flags.MesaCasino = Npclist(MapData(Map, X, Y).NpcIndex).flags.Apostador
        Call SendData(ToIndex, userindex, 0, "ABRU" & UserList(userindex).flags.MesaCasino)
        Exit Sub
    End If
    
    If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
        Call EnviarListaCriaturas(userindex, UserList(userindex).flags.TargetNpc)
        Exit Sub
    End If
    
    
    If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_DM Then
        Call SendData(ToIndex, userindex, 0, "SHWDM")
        Exit Sub
    End If
    'Call SendData(ToIndex, userindex, 0, "SHWDM")
    
    If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_VIEJO Then
        If (UserList(userindex).Stats.ELV >= 40 And UserList(userindex).Stats.RecompensaLevel <= 2) Then
            If Distancia(Npclist(UserList(userindex).flags.TargetNpc).POS, UserList(userindex).POS) > 4 Then
                Call SendData(ToIndex, userindex, 0, "DL")
                Exit Sub
            End If
        End If
        If Not ClaseBase(UserList(userindex).Clase) And Not ClaseTrabajadora(UserList(userindex).Clase) And UserList(userindex).Clase <= GUERRERO Then
            Call SendData(ToIndex, userindex, 0, "RELOM" & UserList(userindex).Clase & "," & UserList(userindex).Stats.RecompensaLevel)
            Exit Sub
        End If
    End If

    If Npclist(MapData(Map, X, Y).NpcIndex).NPCtype = NPCTYPE_NOBLE Then
        If ClaseBase(UserList(userindex).Clase) Or ClaseTrabajadora(UserList(userindex).Clase) Then Exit Sub
    
        If UserList(userindex).Faccion.Bando <> Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion Then
            Call SendData(ToIndex, userindex, 0, Mensajes(Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion, 16) & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
            Exit Sub
        End If
        
        If UserList(userindex).Faccion.Jerarquia = 0 Then
            Call Enlistar(userindex, Npclist(UserList(userindex).flags.TargetNpc).flags.Faccion)
        Else
            Call Recompensado(userindex)
        End If
        
        Exit Sub
    End If
End If

If SegundaPos = 0 Then
If MapData(Map, X, Y).OBJInfo.OBJIndex Then
    UserList(userindex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.OBJIndex
    
    Select Case ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType
        
        Case OBJTYPE_PUERTAS
            Call AccionParaPuerta(Map, X, Y, userindex)
        Case OBJTYPE_CARTELES
            Call AccionParaCartel(Map, X, Y, userindex)
        Case OBJTYPE_FOROS
            Call AccionParaForo(Map, X, Y, userindex)
        Case OBJTYPE_LEÑA
            If MapData(Map, X, Y).OBJInfo.OBJIndex = FOGATA_APAG Then
                Call AccionParaRamita(Map, X, Y, userindex)
            End If
        Case OBJTYPE_ARBOLES
            Call AccionParaArbol(Map, X, Y, userindex)
        
    End Select

ElseIf MapData(Map, X + 1, Y).OBJInfo.OBJIndex Then
    UserList(userindex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.OBJIndex
    Call SendData(ToIndex, userindex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).Name & "," & "OBJ")
    Select Case ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).ObjType
        
        Case 6
            Call AccionParaPuerta(Map, X + 1, Y, userindex)
        
    End Select
ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex Then
    UserList(userindex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex
    Call SendData(ToIndex, userindex, 0, "SELE" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).ObjType & "," & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).Name & "," & "OBJ")
    Select Case ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).ObjType
        
        Case 6
            Call AccionParaPuerta(Map, X + 1, Y + 1, userindex)
        
    End Select
ElseIf MapData(Map, X, Y + 1).OBJInfo.OBJIndex Then
    UserList(userindex).flags.TargetObj = MapData(Map, X, Y + 1).OBJInfo.OBJIndex
    Call SendData(ToIndex, userindex, 0, "SELE" & ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).ObjType & "," & ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).Name & "," & "OBJ")
    Select Case ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).ObjType
        
        Case 6
            Call AccionParaPuerta(Map, X, Y + 1, userindex)
        
    End Select
End If
End If

If SegundaPos = 1 Then
    UserList(userindex).flags.TargetNpc = 0
    UserList(userindex).flags.TargetNpcTipo = 0
    UserList(userindex).flags.TargetUser = 0
    UserList(userindex).flags.TargetObj = 0
Else
'SETIAMOS LAS VARIABLES PARA LA OTRA POS
    SegundaPos = 1
    Map = Map
    X = X
    Y = Y + 1
    GoTo NuevaPos
    '/END
End If






If MapData(Map, X, Y).Agua = 1 Then Call AccionParaAgua(Map, X, Y, userindex)

End Sub
Sub AccionParaRamita(Map As Integer, X As Integer, Y As Integer, userindex As Integer)
On Error Resume Next
Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim raise As Integer, nPos As WorldPos

nPos.Map = Map
nPos.X = X
nPos.Y = Y

If Not MapInfo(UserList(userindex).POS.Map).Pk Then
    Call SendData(ToIndex, userindex, 0, "||No puedes hacer fogatas en zonas seguras." & FONTTYPE_WARNING)
    Exit Sub
End If


If Distancia(nPos, UserList(userindex).POS) > 4 Then
    Call SendData(ToIndex, userindex, 0, "DL")
    Exit Sub
End If

If UserList(userindex).Stats.UserSkills(Supervivencia) > 1 And UserList(userindex).Stats.UserSkills(Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 6 And UserList(userindex).Stats.UserSkills(Supervivencia) <= 10 Then
    Suerte = 2
ElseIf UserList(userindex).Stats.UserSkills(Supervivencia) >= 10 And UserList(userindex).Stats.UserSkills(Supervivencia) Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.OBJIndex = FOGATA
    Obj.Amount = 1
    
    Call SendData(ToIndex, userindex, 0, "7O")
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "FO")
    
    Call MakeObj(ToMap, 0, Map, Obj, Map, X, Y)
    
    Dim Fogatita As New cGarbage
    Fogatita.Map = Map
    Fogatita.X = X
    Fogatita.Y = Y
    Call TrashCollector.Add(Fogatita)
    
Else
    Call SendData(ToIndex, userindex, 0, "8O")
End If


If UserList(userindex).flags.Hambre = 0 And UserList(userindex).flags.Sed = 0 Then
    Call SubirSkill(userindex, Supervivencia)
End If

End Sub
Sub AccionParaAgua(Map As Integer, X As Integer, Y As Integer, userindex As Integer)

If MapData(Map, X, Y).Agua = 0 Then Exit Sub

If UserList(userindex).Stats.UserSkills(Supervivencia) >= 75 And UserList(userindex).Stats.MinAGU < UserList(userindex).Stats.MaxAGU Then
    If UserList(userindex).flags.Muerto Then
        Call SendData(ToIndex, userindex, 0, "MU")
        Exit Sub
    End If
    UserList(userindex).Stats.MinAGU = Minimo(UserList(userindex).Stats.MinAGU + 10, UserList(userindex).Stats.MaxAGU)
    UserList(userindex).flags.Sed = 0
    Call SubirSkill(userindex, Supervivencia, 75)
    Call SendData(ToIndex, userindex, 0, "||Has tomado del agua del mar." & FONTTYPE_INFO)
    Call SendData(ToPCArea, userindex, 0, "TW46")
    Call EnviarHyS(userindex)
End If
    
End Sub
Sub AccionParaArbol(Map As Integer, X As Integer, Y As Integer, userindex As Integer)

If MapData(Map, X, Y).OBJInfo.OBJIndex = 0 Then Exit Sub
If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType <> OBJTYPE_ARBOLES Then Exit Sub

If UserList(userindex).Stats.UserSkills(Supervivencia) >= 85 And UserList(userindex).Stats.MinHam < UserList(userindex).Stats.MaxHam Then
    If UserList(userindex).flags.Muerto Then
        Call SendData(ToIndex, userindex, 0, "MU")
        Exit Sub
    End If
    UserList(userindex).Stats.MinHam = Minimo(UserList(userindex).Stats.MinHam + 10, UserList(userindex).Stats.MaxHam)
    UserList(userindex).flags.Hambre = 0
    Call SubirSkill(userindex, Supervivencia, 75)
    Call SendData(ToIndex, userindex, 0, "||Has comido de los frutos del árbol." & FONTTYPE_INFO)
    Call SendData(ToPCArea, userindex, 0, "TW7")
    Call EnviarHyS(userindex)
End If

End Sub
Sub AccionParaForo(Map As Integer, X As Integer, Y As Integer, userindex As Integer)
On Error Resume Next


Dim f As String, tit As String, men As String, Base As String, auxcad As String
f = App.Path & "\foros\" & UCase$(ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ForoID) & ".for"
If FileExist(f, vbNormal) Then
    Dim Num As Integer
    Num = val(GetVar(f, "INFO", "CantMSG"))
    Base = Left$(f, Len(f) - 4)
    Dim i As Integer
    Dim N As Integer
    For i = 1 To Num
        N = FreeFile
        f = Base & i & ".for"
        Open f For Input Shared As #N
        Input #N, tit
        men = ""
        auxcad = ""
        Do While Not EOF(N)
            Input #N, auxcad
            men = men & vbCrLf & auxcad
        Loop
        Close #N
        Call SendData(ToIndex, userindex, 0, "FMSG" & tit & Chr$(176) & men)
        
    Next
End If
Call SendData(ToIndex, userindex, 0, "MFOR")
End Sub


Sub AccionParaPuerta(Map As Integer, X As Integer, Y As Integer, userindex As Integer)
On Error Resume Next

Dim MiObj As Obj
Dim wp As WorldPos

If Not (Distance(UserList(userindex).POS.X, UserList(userindex).POS.Y, X, Y) > 2) Then
    If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Llave = 0 Then
        If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Cerrada Then
                
                If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Llave = 0 Then
                          
                     MapData(Map, X, Y).OBJInfo.OBJIndex = ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).IndexAbierta
                                  
                     Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                     
                     
                     MapData(Map, X, Y).Blocked = 0
                     MapData(Map, X - 1, Y).Blocked = 0
                     
                     
                     Call Bloquear(ToMap, 0, Map, Map, X, Y, 0)
                     Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 0)
                     
                       
                     
                     SendData ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_PUERTA
                    
                Else
                     Call SendData(ToIndex, userindex, 0, "9O")
                End If
        Else
                
                MapData(Map, X, Y).OBJInfo.OBJIndex = ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).IndexCerrada
                
                Call MakeObj(ToMap, 0, Map, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                
                
                MapData(Map, X, Y).Blocked = 1
                MapData(Map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(ToMap, 0, Map, Map, X - 1, Y, 1)
                Call Bloquear(ToMap, 0, Map, Map, X, Y, 1)
                
                SendData ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_PUERTA
        End If
        
        UserList(userindex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.OBJIndex
    Else
        Call SendData(ToIndex, userindex, 0, "9O")
    
    End If
Else
    Call SendData(ToIndex, userindex, 0, "DL")
End If

End Sub
Sub AccionParaCartel(Map As Integer, X As Integer, Y As Integer, userindex As Integer)
On Error Resume Next

Dim MiObj As Obj

If ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = 8 Then
  
  If Len(ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Texto) > 0 Then
       Call SendData(ToIndex, userindex, 0, "MCAR" & _
        ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Texto & _
        Chr$(176) & ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).GrhSecundario)
  End If
  
End If

End Sub

Sub AccionRespuestaGm(ByVal Nick As String, ByVal Texto As String)
    Dim Indice As Integer
    Indice = NameIndex(Nick)
    Call SendData(ToIndex, Indice, 0, "SS" & Texto)
End Sub


