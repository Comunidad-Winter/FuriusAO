Attribute VB_Name = "CaptureTheFlag"
Public Const MAP_CTF As Integer = 194 '196
Public Const MAP_CTC As Integer = 196 '196
Public Const BANDERAINDEXCRIMI As Integer = 997
Public Const BANDERAINDEXCIUDA As Integer = 996

Public Recaudado As Long

Public BanderaCrimiPos As WorldPos
Public BanderaCiudaPos As WorldPos


Public Sub DevolverBandera(Bando As Integer, UserNombre As String)
Dim ASD As Obj
ASD.Amount = 1
If Bando = 2 Then ' criminalk
ASD.OBJIndex = BANDERAINDEXCRIMI
    If UCase$(UserNombre) <> "NO4NAME" Then
    Call SendData(ToMap, 0, MAP_CTF, "||Capture the Flag> La bandera Criminal fue devuelta por " & UserNombre & FONTTYPE_CAOS)
    Call SendData(ToMap, 0, MAP_CTC, "||Capture the Flag> La bandera Criminal fue devuelta por " & UserNombre & FONTTYPE_CAOS)
    End If
    Call MakeObj(ToMap, 0, MAP_CTF, ASD, BanderaCrimiPos.Map, BanderaCrimiPos.x, BanderaCrimiPos.Y)

ElseIf Bando = 1 Then 'ciuda
ASD.OBJIndex = BANDERAINDEXCIUDA
    If UCase$(UserNombre) <> "NO4NAME" Then
    Call SendData(ToMap, 0, MAP_CTC, "||Capture the Flag> La bandera Real fue devuelta por " & UserNombre & FONTTYPE_ARMADA)
    Call SendData(ToMap, 0, MAP_CTF, "||Capture the Flag> La bandera Real fue devuelta por " & UserNombre & FONTTYPE_ARMADA)
    End If
    Call MakeObj(ToMap, 0, MAP_CTC, ASD, BanderaCiudaPos.Map, BanderaCiudaPos.x, BanderaCiudaPos.Y)
End If
End Sub


Sub Gano(userindex As Integer)
    Call SumaPuntos(userindex, 25)
    Select Case UserList(userindex).Faccion.Bando
    
    Case Real
    If MapData(BanderaCiudaPos.Map, BanderaCiudaPos.x, BanderaCiudaPos.Y).OBJInfo.OBJIndex = BANDERAINDEXCIUDA Then
    Call DevolverBandera(Real, "no4name")
    Call DevolverBandera(Caos, "no4name")
    Call SendData(ToAll, 0, 0, "||Gana equipo azul" & FONTTYPE_ARMADA)
    Else
    Exit Sub
    End If
    
    Case Caos
    If MapData(BanderaCrimiPos.Map, BanderaCrimiPos.x, BanderaCrimiPos.Y).OBJInfo.OBJIndex = BANDERAINDEXCRIMI Then
    Call DevolverBandera(Real, "no4name")
    Call DevolverBandera(Caos, "no4name")
    Call SendData(ToAll, 0, 0, "||Gana equipo rojo" & FONTTYPE_CAOS)
    Else
    Exit Sub
    End If
       
    End Select
    
    Dim LoopC As Integer

    For LoopC = 1 To LastUser
        If UserList(LoopC).POS.Map = MAP_CTC Or UserList(LoopC).POS.Map = MAP_CTF Then
            If UserList(LoopC).Faccion.Bando = Real Then
            Call WarpUserChar(LoopC, 206, 12, 20, True)
            Else
            Call WarpUserChar(LoopC, 206, 59, 73, True)
            End If
        End If
    Next
    
    
Dim i As Integer
Dim Slotx As Integer
For i = 1 To LastUser
Slotx = 1
    If TieneObjetos(BANDERAINDEXCRIMI, 1, i) Then
        Do Until UserList(i).Invent.Object(Slotx).OBJIndex = BANDERAINDEXCRIMI
        Slotx = Slotx + 1
        Loop
        Call QuitarVariosItem(userindex, Slotx, 1)
 '       Call UpdateUserInv(True, i, val(Slotx))
        Slotx = 1
    End If
    
    If TieneObjetos(BANDERAINDEXCIUDA, 1, i) Then
        Do Until UserList(i).Invent.Object(Slotx).OBJIndex = BANDERAINDEXCIUDA
        Slotx = Slotx + 1
        Loop
        Call QuitarVariosItem(i, Slotx, 1)
'        Call UpdateUserInv(True, i, val(Slotx))
    End If

Next i
End Sub


Public Sub PagarC(Bando As Byte)
Dim i As Integer
'CPts
'0 , 1 y 2 bando
Dim TotalPts As Long
For i = 1 To LastUser
If UserList(i).Stats.CPts > 0 And UserList(i).Faccion.Bando = Bando Then TotalPts = UserList(i).Stats.CPts + TotalPts
Next i
i = 0
Dim Ganancia As Long
For i = 1 To LastUser
    If UserList(i).Stats.CPts > 0 And UserList(i).Faccion.Bando = Bando Then
    Ganancia = Int(UserList(i).Stats.CPts / TotalPts * Recaudado)
    UserList(i).Stats.GLD = UserList(i).Stats.GLD + Ganancia
    Call SendUserORO(i)
    End If
Next i
Call SendData(ToAdmins, 0, 0, "||Premios de Capture the Flag repartidos" & FONTTYPE_VENENO)
End Sub

Public Sub SumaPuntos(userindex As Integer, Cantidad As Byte)
UserList(userindex).Stats.CPts = UserList(userindex).Stats.CPts + Cantidad
End Sub
