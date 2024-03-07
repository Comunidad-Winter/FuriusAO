Attribute VB_Name = "ModMascotas"

Public Function TransformarMascota(ByVal CBM As Integer, ByVal userindex As Integer)
Dim NombreCriatura As String
Dim ObjFile As String
Dim GemaMascota As Obj
Dim GrhIndexMIO As Integer
Dim Nivel As Integer


NumObjDatas = val(GetVar(DatPath & "Mascotas.dat", "INIT", "NumMascotas"))
ObjFile = DatPath & "Mascotas.dat"

Dim Nombre As String
If Npclist(CBM).Name = "Elefante" Then
     
        MinDefMag = 1
        MaxDefMag = 5
        MinDef = 1
        MaxDef = 4
        MaxHIT = 5
        MinHIT = 1
        MaxHITMag = 6
        MinHITMag = 2
        NumRopaje = 282
        GraficoIndex = 705

ElseIf Npclist(CBM).Name = NPC2 Then



Else
Exit Function
End If
Nombre = Npclist(CBM).Name
Call SendData(ToIndex, userindex, 0, "||Has capturado a la criatura... Aguarda unos segundos y la verás aparecer en el inventario." & FONTTYPE_INFO)
NumObjDatas = NumObjDatas + 1
Call MuereNpc(CBM, 0)

Call WriteVar(ObjFile, "INIT", "NumMascotas", val(NumObjDatas))
Call WriteVar(ObjFile, "M" & NumObjDatas, "Name", (Nombre & NumObjDatas))
Call WriteVar(ObjFile, "M" & NumObjDatas, "Alias", Nombre)
Call WriteVar(ObjFile, "M" & NumObjDatas, "Level", 1)
Call WriteVar(ObjFile, "M" & NumObjDatas, "MinDef", val(MinDef))
Call WriteVar(ObjFile, "M" & NumObjDatas, "MaxDef", val(MaxDef))
Call WriteVar(ObjFile, "M" & NumObjDatas, "MinDefMag", val(MinDefMag))
Call WriteVar(ObjFile, "M" & NumObjDatas, "MaxDefMag", val(MaxDefMag))
Call WriteVar(ObjFile, "M" & NumObjDatas, "MinHIT", val(MinHIT))
Call WriteVar(ObjFile, "M" & NumObjDatas, "MinHITMag", val(MinHITMag))
Call WriteVar(ObjFile, "M" & NumObjDatas, "MaxHITMag", val(MaxHITMag))

Call WriteVar(ObjFile, "M" & NumObjDatas, "MaxHIT", val(MaxHIT))
Call WriteVar(ObjFile, "M" & NumObjDatas, "OBJTYPE", 200)
Call WriteVar(ObjFile, "M" & NumObjDatas, "Numropaje", val(NumRopaje))
Call WriteVar(ObjFile, "M" & NumObjDatas, "GrhIndex", val(GraficoIndex))

Call WriteVar(ObjFile, "M" & NumObjDatas, "ObjetoMascota", val(NumObjDatas))
Call WriteVar(ObjFile, "M" & NumObjDatas, "MinExp", 0)
Call WriteVar(ObjFile, "M" & NumObjDatas, "MaxExp", 500)
Call WriteVar(ObjFile, "M" & NumObjDatas, "NoSeCae", 1)
Dim ObjsActuales As Integer
ObjsActuales = GetVar(DatPath & "obj.dat", "INIT", "NumOBJs") + 2
ReDim Preserve ObjData(0 To ObjsActuales + NumObjDatas) As ObjData
With ObjData(ObjsActuales + NumObjDatas)
.Name = Nombre & NumObjDatas
.MinHIT = MinHIT
.MaxHIT = MaxHIT
.ObjetoMascota = NumObjDatas
.level = 1
.MinExp = 0
.MaxExp = 500
.Alias = Nombre
.MaxDef = MaxDef
.MinDef = MinDef
.MinDefMag = MinDefMag
.MaxDefMag = MaxDefMag
.MinHITMag = MinHITMag
.MaxHITMag = MaxHITMag
.Ropaje = NumRopaje
.ObjType = OBJTYPE_MASCOTA
.GrhIndex = GraficoIndex
.NoSeCae = True
End With
Call SendData(ToIndex, userindex, 0, "||Tu nueva mascota es un " & ObjData(ObjsActuales + NumObjDatas).Name & FONTTYPE_BLANCO)
GemaMascota.Amount = 1
GemaMascota.OBJIndex = NumObjDatas + ObjsActuales
TransformarMascota = True
If Not MeterItemEnInventario(userindex, GemaMascota) Then Call TirarItemAlPiso(UserList(userindex).POS, GemaMascota)
End Function
Sub MascotaSubirExp(ByVal userindex As Integer)
Dim ObjIndexx As Integer
ObjIndexx = UserList(userindex).Invent.MascotaEqpObjIndex
Dim AumentoMinHit As Integer
Dim AumentoMaxHit As Integer
Dim AumentoMinDef As Integer
Dim AumentoMaxDef As Integer
Dim AumentoMinDefMag As Integer
Dim AumentoMaxDefMag As Integer
Dim AumentoMinHitMag As Integer
Dim AumentoMaxHitMag As Integer


If ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinExp >= ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxExp * 3 Then
    
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SOUND_NIVEL)
    Call SendData(ToIndex, userindex, 0, "||Tu mascota ha subido de nivel!" & vbCrLf & "Su nuevo nivel es " & val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).level) + 1 & FONTTYPE_INFO)
    
    
    ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxExp = (CInt(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxExp / ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).level)) * CInt(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxExp * level + 1)
    ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).level = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).level + 1
    ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinExp = 0
    Select Case ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).Name
        Case "Elefante"
        AumentoMinHit = 1
        AumentoMaxHit = 5
        AumentoMinDef = 1
        AumentoMaxDef = 5
        AumentoMinDefMag = 1
        AumentoMaxDefMag = 5
        AumentoMinHitMag = 2
        AumentoMaxHitMag = 6
        Case Else
        AumentoMinHit = 1
        AumentoMaxHit = 5
        AumentoMinDef = 1
        AumentoMaxDef = 5
        AumentoMinDefMag = 1
        AumentoMaxDefMag = 5
        AumentoMinHitMag = 2
        AumentoMaxHitMag = 6
    End Select
   
   
   
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHIT = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHIT + AumentoMinHit
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHIT = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHIT + AumentoMaxHit
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinDef = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinDef + AumentoMinDef
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxDef = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxDef + AumentoMaxDef
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinDefMag = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinDefMag + AumentoMinDefMag
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxDefMag = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxDefMag + AumentoMaxDefMag
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHITMag = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHITMag + AumentoMinHitMag
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHITMag = ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHITMag + AumentoMaxHitMag

Call WriteVar(DatPath & "Mascotas.dat", "M" & ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota, "MaxExp", ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxExp * ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).level)
Call WriteVar(DatPath & "Mascotas.dat", "M" & ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota, "MinExp", 0)
Call WriteVar(DatPath & "Mascotas.dat", "M" & ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota, "Level", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).level))
Call WriteVar(DatPath & "Mascotas.dat", "M" & val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota), "MinHit", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHIT))
Call WriteVar(DatPath & "Mascotas.dat", "M" & ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota, "MaxHit", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHIT))
Call WriteVar(DatPath & "Mascotas.dat", "M" & val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota), "MinDef", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinDef))
Call WriteVar(DatPath & "Mascotas.dat", "M" & ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota, "MaxDef", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxDef))
Call WriteVar(DatPath & "Mascotas.dat", "M" & val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota), "MinDefMag", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinDefMag))
Call WriteVar(DatPath & "Mascotas.dat", "M" & ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota, "MaxDefMag", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxDefMag))
Call WriteVar(DatPath & "Mascotas.dat", "M" & ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota, "MinHITMag", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHITMag))
Call WriteVar(DatPath & "Mascotas.dat", "M" & ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota, "MaxHITMag", val(ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHITMag))
   
   
    SendData ToIndex, userindex, 0, "||El golpe minimo de tu mascota ha aumentado en " & AumentoMinHit & " puntos." & FONTTYPE_INFO
    SendData ToIndex, userindex, 0, "||El golpe maximo de tu mascota ha aumentado en " & AumentoMaxHit & " puntos." & FONTTYPE_INFO
    SendData ToIndex, userindex, 0, "||La defensa minima de tu mascota ha aumentado en " & AumentoMinDef & " puntos." & FONTTYPE_INFO
    SendData ToIndex, userindex, 0, "||La defensa maxima de tu mascota ha aumentado en " & AumentoMaxDef & " puntos." & FONTTYPE_INFO
    SendData ToIndex, userindex, 0, "||La defensa mágica minima de tu mascota ha aumentado en " & AumentoMinDefMag & " puntos." & FONTTYPE_INFO
    SendData ToIndex, userindex, 0, "||La defensa mágica maxima de tu mascota ha aumentado en  " & AumentoMaxDefMag & " puntos." & FONTTYPE_INFO
    SendData ToIndex, userindex, 0, "||El golpe mágico minimo de tu mascota ha aumentado en " & AumentoMinHitMag & " puntos." & FONTTYPE_INFO
    SendData ToIndex, userindex, 0, "||El golpe mágico maximo de tu mascota ha aumentado en " & AumentoMaxHitMag & " puntos." & FONTTYPE_INFO
    
    Call LoadOBJData
    'Call MascotaSubirExp(UserIndex)
    AumentoMinHit = 0
    AumentoMaxHit = 0
    AumentoMinDef = 0
    AumentoMaxDef = 0
    AumentoMinDefMag = 0
    AumentoMaxDefMag = 0
    AumentoMinHitMag = 0
    AumentoMaxHitMag = 0
End If


End Sub
Sub SendMascBox(ByVal userindex As Integer)
If UserList(userindex).flags.Montado = 1 Then
Call SendData(ToIndex, userindex, 0, "MON" & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).Alias & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).ObjetoMascota & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHIT & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHIT & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).level & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).Name & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinDef & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxDef & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinDefMag & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxDefMag & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinHITMag & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MaxHITMag & "," & _
ObjData(UserList(userindex).Invent.MascotaEqpObjIndex).MinExp)
End If
End Sub

