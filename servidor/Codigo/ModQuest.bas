Attribute VB_Name = "ModQuest"
Public Const MAXQUESTS = 100
Public Type tQuest
    Tipo As Integer
    Tiempo As Integer
    Usuarios As Integer
    Map As Integer
    NPCs As Integer
    BuscaNPC As String
    MataNPC As String
    Oro As Integer
    Obj As Integer
    Cant As Integer
End Type
Public Quest(1 To MAXQUESTS) As tQuest
Public NUMQUESTS As Integer
Public Sub CargarQuests()
Dim LoopC As Integer
NUMQUESTS = val(GetVar(DatPath & "QUESTS.dat", "INIT", "NUM"))
For LoopC = 1 To NUMQUESTS
    Quest(LoopC).Tipo = val(GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "Tipo"))
    Quest(LoopC).Tiempo = val(GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "Tiempo"))
    Quest(LoopC).Usuarios = val(GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "Usuarios"))
    Quest(LoopC).Map = val(GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "Map"))
    Quest(LoopC).NPCs = val(GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "NPCs"))
    Quest(LoopC).BuscaNPC = (GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "BuscaNpc"))
    Quest(LoopC).MataNPC = (GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "MataNpc"))
    Quest(LoopC).Oro = (GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "Oro"))
    Quest(LoopC).Obj = (GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "Obj"))
    Quest(LoopC).Cant = (GetVar(DatPath & "QUESTS.dat", "QUEST" & LoopC, "Cant"))
Next LoopC
End Sub
Public Sub RealizarQuest(UserIndex As Integer)
Dim AzarQuest As Integer
Dim Resolucion As String
Dim PosAzar As WorldPos
AzarQuest = RandomNumber(1, NUMQUESTS)
Select Case Quest(AzarQuest).Tipo
    Case 1
        Resolucion = "Debes ir a el mapa " & Quest(AzarQuest).Map & " en menos de " & Quest(AzarQuest).Tiempo & " minutos para recibir tu recompensa. PISTA: Puedes guiarte con el mapa del server apretando la tecla Q."
    Case 2
        Resolucion = "Debes ir a el mapa " & Quest(AzarQuest).Map & " en menos de " & Quest(AzarQuest).Tiempo & " minutos y pedirle a alguien que este ahi que te clicke y escriba /LOLOGRO para recibir tu recompensa. PISTA: Puedes guiarte con el mapa del server apretando la tecla Q."
    Case 3
        Resolucion = "Mata " & Quest(AzarQuest).NPCs & " bichos en menos de " & Quest(AzarQuest).Tiempo & " minutos para recibir tu recompensa."
    Case 4
        If Criminal(UserIndex) Then
            Resolucion = "Mata " & Quest(AzarQuest).Usuarios & " ciudadanos en menos de " & Quest(AzarQuest).Tiempo & " minutos para recibir tu recompensa."
        Else
            Resolucion = "Mata " & Quest(AzarQuest).Usuarios & " criminales en menos de " & Quest(AzarQuest).Tiempo & " minutos para recibir tu recompensa."
        End If
    Case 5
        Resolucion = "Encuentra al NPC '" & Quest(AzarQuest).BuscaNPC & "' y tipea /LOLOGRO, en menos de " & Quest(AzarQuest).Tiempo & " minutos para recibir tu recompensa. PISTA: Este npc se puede encontrar en una ciudad."
        Call PosiblesLugar(RandomCity)
        PosAzar = PosiblesLugarz
        If UserList(UserIndex).flags.Privilegios = 3 Then Call SendData(ToIndex, UserIndex, 0, "||" & PosAzar.Map & " " & PosAzar.X & " " & PosAzar.Y & FONTTYPE_INFO)
        
        UserList(UserIndex).flags.NpcQuest = SpawnNpc(127, PosAzar, True, False)
        Npclist(UserList(UserIndex).flags.NpcQuest).Name = Quest(AzarQuest).BuscaNPC
    Case 6
        Resolucion = "Mata al NPC '" & Quest(AzarQuest).MataNPC & "' en menos de " & Quest(AzarQuest).Tiempo & " minutos para recibir tu recompensa."
        Call PosiblesLugar(RandomCity)
        PosAzar = PosiblesLugarz
        If UserList(UserIndex).flags.Privilegios = 3 Then Call SendData(ToIndex, UserIndex, 0, "||" & PosAzar.Map & " " & PosAzar.X & " " & PosAzar.Y & FONTTYPE_INFO)

        UserList(UserIndex).flags.NpcQuest = SpawnNpc(128, PosAzar, True, False)
        Npclist(UserList(UserIndex).flags.NpcQuest).Name = Quest(AzarQuest).MataNPC
End Select
Call SendData(ToIndex, UserIndex, UserList(UserIndex).POS.Map, "||" & vbWhite & "°" & Resolucion & "°" & str(Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex))
UserList(UserIndex).flags.EnQuest = True
UserList(UserIndex).flags.Quest = AzarQuest
UserList(UserIndex).flags.QuestR.Tiempo = Quest(AzarQuest).Tiempo
Call ActualizarQuestInfo(UserIndex)
End Sub
Public Sub ActualizarQuestInfo(UserIndex As Integer)
If UserList(UserIndex).flags.Quest = 0 Then Exit Sub
Select Case Quest(UserList(UserIndex).flags.Quest).Tipo
    Case 1
        Call SendData(ToIndex, UserIndex, 0, "QMSGTiempo restante " & UserList(UserIndex).flags.QuestR.Tiempo & " minutos. Ve rapido al mapa " & Quest(UserList(UserIndex).flags.Quest).Map & " antes de que termine el tiempo.")
    Case 2
        Call SendData(ToIndex, UserIndex, 0, "QMSGTiempo restante " & UserList(UserIndex).flags.QuestR.Tiempo & " minutos. Ve rapido al mapa " & Quest(UserList(UserIndex).flags.Quest).Map & " antes de que termine el tiempo y pidele a alguien que te ponga /LOLOGRO.")
    Case 3
        Call SendData(ToIndex, UserIndex, 0, "QMSGTiempo restante " & UserList(UserIndex).flags.QuestR.Tiempo & " minutos. Has Matado " & UserList(UserIndex).flags.QuestR.NPCs & "/" & Quest(UserList(UserIndex).flags.Quest).NPCs & " bichos.")
    Case 4
        Call SendData(ToIndex, UserIndex, 0, "QMSGTiempo restante " & UserList(UserIndex).flags.QuestR.Tiempo & " minutos. Has Matado " & UserList(UserIndex).flags.QuestR.Usuarios & "/" & Quest(UserList(UserIndex).flags.Quest).Usuarios & " usuarios.")
    Case 5
        Call SendData(ToIndex, UserIndex, 0, "QMSGTiempo restante " & UserList(UserIndex).flags.QuestR.Tiempo & " minutos. Recorre las ciudades hasta encontrar al npc perdido antes de que termine el tiempo y cuando lo encuentres pon /LOLOGRO")
    Case 6
        Call SendData(ToIndex, UserIndex, 0, "QMSGTiempo restante " & UserList(UserIndex).flags.QuestR.Tiempo & " minutos. Recorre las ciudades hasta encontrar el npc enemigo y aniquilalo antes de que se termine el tiempo.")
End Select
End Sub
Public Sub RebisarNPCsMatados(UserIndex As Integer)
Call SendData(ToIndex, UserIndex, 0, "QMSGTiempo restante " & UserList(UserIndex).flags.QuestR.Tiempo & " minutos. Has Matado " & UserList(UserIndex).flags.QuestR.NPCs & "/" & Quest(UserList(UserIndex).flags.Quest).NPCs & " bichos.")
If UserList(UserIndex).flags.QuestR.NPCs >= Quest(UserList(UserIndex).flags.Quest).NPCs Then
    Call SendData(ToIndex, UserIndex, 0, "||Quest completada. Ve con el organizador para recibir tu recompensa" & FONTTYPE_INFO)
    Call LogroLaQuest(UserIndex)
End If
End Sub
Public Sub RebisarUsuariosMatados(UserIndex As Integer)
Call SendData(ToIndex, UserIndex, 0, "QMSGTiempo restante " & UserList(UserIndex).flags.QuestR.Tiempo & " minutos. Has Matado " & UserList(UserIndex).flags.QuestR.Usuarios & "/" & Quest(UserList(UserIndex).flags.Quest).Usuarios & " usuarios.")
If UserList(UserIndex).flags.QuestR.Usuarios >= Quest(UserList(UserIndex).flags.Quest).Usuarios Then
    Call SendData(ToIndex, UserIndex, 0, "||Quest completada. Ve con el organizador para recibir tu recompensa" & FONTTYPE_INFO)
    Call LogroLaQuest(UserIndex)
End If
End Sub
Public Sub LogroLaQuest(UserIndex As Integer)
UserList(UserIndex).flags.QuestCumplida = True
Call SendData(ToIndex, UserIndex, 0, "QMSGQUEST COMPLETADA. Ahora vuelve para recibir tu recompensa.")
UserList(UserIndex).flags.EnQuest = False
End Sub
Public Sub ResetQuestInfo(UserIndex As Integer)
UserList(UserIndex).flags.EnQuest = False
UserList(UserIndex).flags.QuestCumplida = False
UserList(UserIndex).flags.Quest = 0
UserList(UserIndex).flags.QuestR.Tiempo = 0
UserList(UserIndex).flags.QuestR.Tipo = 0
UserList(UserIndex).flags.QuestR.Usuarios = 0
UserList(UserIndex).flags.QuestR.Map = 0
UserList(UserIndex).flags.QuestR.NPCs = 0
UserList(UserIndex).flags.QuestR.BuscaNPC = ""
UserList(UserIndex).flags.QuestR.MataNPC = ""
UserList(UserIndex).flags.NpcQuest = 0
UserList(UserIndex).TiempoQuest = 0
End Sub
Public Sub RecompensaQuest(UserIndex As Integer)
If Quest(UserList(UserIndex).flags.Quest).Oro > 0 Then
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Quest(UserList(UserIndex).flags.Quest).Oro
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "¡¡¡Felicitaciones!!! Has ganado " & Quest(UserList(UserIndex).flags.Quest).Oro & " monedas de oro.°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
    Call SendUserStatsBox(val(UserIndex))
End If
If Quest(UserList(UserIndex).flags.Quest).Obj > 0 Then
    Dim MiObj2 As Obj
    MiObj2.Amount = Quest(UserList(UserIndex).flags.Quest).Cant
    MiObj2.OBJIndex = Quest(UserList(UserIndex).flags.Quest).Obj
    If Not MeterItemEnInventario(UserIndex, MiObj2) Then
        Call TirarItemAlPiso(UserList(UserIndex).POS, MiObj2)
    End If
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "¡¡¡Felicitaciones!!! Has ganado " & Quest(UserList(UserIndex).flags.Quest).Cant & " " & ObjData(Quest(UserList(UserIndex).flags.Quest).Obj).Name & ".°" & Npclist(UserList(UserIndex).flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
End If
If Quest(UserList(UserIndex).flags.Quest).Tipo = 5 Then
    Call QuitarNPC(UserList(UserIndex).flags.NpcQuest)
End If
If UserList(UserIndex).flags.QuestCumplida = False Then
    If Quest(UserList(UserIndex).flags.Quest).Tipo = 6 Then
        Call QuitarNPC(UserList(UserIndex).flags.NpcQuest)
    End If
End If
Call SendData(ToIndex, UserIndex, 0, "QMSG ")
Call ResetQuestInfo(UserIndex)
UserList(UserIndex).TiempoQuest = 60
End Sub
Public Function RandomCity() As Integer
Dim AzarN As Integer
AzarN = RandomNumber(1, 6)
Select Case AzarN
    Case 1
        RandomCity = 1
    Case 2
        RandomCity = 34
    Case 3
        RandomCity = 59
    Case 4
        RandomCity = 60
    Case 5
        RandomCity = 81
    Case 6
        RandomCity = 83
End Select
End Function
Public Sub PosiblesLugar(Mapa As Integer)
Dim AzarN As Integer
Dim Xp As Integer
Dim Yp As Integer
AzarN = RandomNumber(1, 3)
Select Case Mapa
    Case 1
        Select Case AzarN
            Case 1
                Xp = 34
                Yp = 58
            Case 2
                Xp = 25
                Yp = 77
            Case 3
                Xp = 63
                Yp = 59
        End Select
    Case 34
        Select Case AzarN
            Case 1
                Xp = 41
                Yp = 61
            Case 2
                Xp = 84
                Yp = 38
            Case 3
                Xp = 19
                Yp = 85
        End Select
    Case 59
        Select Case AzarN
            Case 1
                Xp = 21
                Yp = 76
            Case 2
                Xp = 44
                Yp = 64
            Case 3
                Xp = 63
                Yp = 44
        End Select
    Case 60
        Select Case AzarN
            Case 1
                Xp = 73
                Yp = 50
            Case 2
                Xp = 53
                Yp = 51
            Case 3
                Xp = 63
                Yp = 90
        End Select
    Case 81
        Select Case AzarN
            Case 1
                Xp = 34
                Yp = 44
            Case 2
                Xp = 45
                Yp = 81
            Case 3
                Xp = 69
                Yp = 62
        End Select
    Case 83
        Select Case AzarN
            Case 1
                Xp = 39
                Yp = 71
            Case 2
                Xp = 65
                Yp = 53
            Case 3
                Xp = 78
                Yp = 37
        End Select
    Case Else
        Xp = 50
        Yp = 50
End Select
PosiblesLugarz.X = Xp
PosiblesLugarz.Y = Yp
PosiblesLugarz.Map = Mapa
End Sub
