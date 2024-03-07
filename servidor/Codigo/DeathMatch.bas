Attribute VB_Name = "DeathMatch"
'CARGAR EL INI CON CALL CARGARINI
'Un comando que sea "Call CargarClase(userindex,rdata)"
'Y el user deberia llevar el flag EnDM


'MODULO DEATHMATCH SIMLPE
Public RutaDIni As String

Const DMComida As Integer = 26 ' pollo
Const DMBebida As Integer = 42 ' vino


Type ObjDMT
    Num As Integer
    Cant As Integer
    Equipado As Byte
End Type

Type ClaseDM
    NClase As Integer
    Nombre As String
    Vida As Integer
    Mana As Integer
    MinG As Integer
    MaxG As Integer
    ObjC As Integer
    ObjDM() As ObjDMT
    Raza As String
    Recom(1 To 3) As Byte
    Hechizos(1 To MAXUSERHECHIZOS) As Integer
End Type

Public ClaseDM() As ClaseDM

'CARACTERISTICAS DEL DEATHMATCH
Public OroEntrada As Integer
Public OroMuerte As Integer
Public OroKill As Integer
Public MapaDM As Integer
Public MapaEQ As Integer
Public MapaTa As Integer
Public MapaTb As Integer
Public DMActivado As Integer
Public DM_MBienvenida As String
Public DM_MMuerte As String
Public DM_MKill As String
Public DM_MinLVL As Integer
'/CARACTERISTICAS DEL DEATHMATCH


Public Sub CargarIniDM()
Dim i As Integer
RutaDIni = App.Path & "\DeathMatch.ini"
'CARGAMOS LO PRINCIPAL QUE ES EL ORO ETC.
DMActivado = val(GetVar(RutaDIni, "GENERAL", "Activado"))
OroEntrada = val(GetVar(RutaDIni, "GENERAL", "OroEntrada"))
OroMuerte = val(GetVar(RutaDIni, "GENERAL", "OroMuerte"))
OroKill = val(GetVar(RutaDIni, "GENERAL", "OroKill"))
MapaDM = val(GetVar(RutaDIni, "GENERAL", "Mapa"))
MapaEQ = val(GetVar(RutaDIni, "GENERAL", "MapaEquipamiento"))
MapaTa = val(GetVar(RutaDIni, "GENERAL", "MapaS1"))
MapaTb = val(GetVar(RutaDIni, "GENERAL", "MapaS2"))
DM_MinLVL = val(GetVar(RutaDIni, "GENERAL", "NivelMinimo"))
DM_MBienvenida = GetVar(RutaDIni, "GENERAL", "MBienvenida") & " "
DM_MMuerte = GetVar(RutaDIni, "GENERAL", "MMuerte") & " "
DM_MKill = GetVar(RutaDIni, "GENERAL", "MKill") & " "

'loaded(?)


'Cargamos las clases
Dim CntClases As Integer
Dim ClaseActual As String
Dim H As Integer

CntClases = val(GetVar(RutaDIni, "GENERAL", "Clases"))

If CntClases = 0 Then Exit Sub
ReDim ClaseDM(CntClases)



For i = 1 To CntClases
ClaseActual = GetVar(RutaDIni, "GENERAL", "Clase" & i)
ClaseDM(i).Nombre = ClaseActual
ClaseDM(i).NClase = i
ClaseDM(i).Mana = val(GetVar(RutaDIni, ClaseActual, "Mana"))
ClaseDM(i).Vida = val(GetVar(RutaDIni, ClaseActual, "Vida"))
ClaseDM(i).MaxG = val(GetVar(RutaDIni, ClaseActual, "MaxG"))
ClaseDM(i).MinG = val(GetVar(RutaDIni, ClaseActual, "MinG"))
ClaseDM(i).ObjC = val(GetVar(RutaDIni, ClaseActual, "Objs"))
ClaseDM(i).Raza = GetVar(RutaDIni, ClaseActual, "Raza")
ClaseDM(i).Recom(1) = GetVar(RutaDIni, ClaseActual, "Recom1")
ClaseDM(i).Recom(2) = GetVar(RutaDIni, ClaseActual, "Recom2")
ClaseDM(i).Recom(3) = GetVar(RutaDIni, ClaseActual, "Recom3")




ReDim ClaseDM(i).ObjDM(ClaseDM(i).ObjC)

For H = 1 To ClaseDM(i).ObjC
ClaseDM(i).ObjDM(H).Num = ReadField$(1, (GetVar(RutaDIni, ClaseActual, "Obj" & H)), Asc(","))
ClaseDM(i).ObjDM(H).Cant = ReadField$(2, (GetVar(RutaDIni, ClaseActual, "Obj" & H)), Asc(","))
ClaseDM(i).ObjDM(H).Equipado = val(ReadField$(3, (GetVar(RutaDIni, ClaseActual, "Obj" & H)), Asc(",")))
DoEvents
Next H



For H = 1 To MAXUSERHECHIZOS
ClaseDM(i).Hechizos(H) = val(GetVar(RutaDIni, ClaseActual, "H" & H))
DoEvents
Next H


DoEvents
Next i


End Sub


Public Sub CargarClase(userindex As Integer, Clase As String)
Dim NoSave As Boolean
Dim i As Integer
Dim x As Integer
If DMActivado = 0 Then
Call SendData(ToIndex, userindex, 0, "||Disculpe, el DeathMatch está desactivado en este momento." & FONTTYPE_BLANCO)
Exit Sub
End If

If UserList(userindex).Stats.ELV < DM_MinLVL Then
Call SendData(ToIndex, userindex, 0, "||No posees el nivel suficiente." & FONTTYPE_BLANCO)
Exit Sub
End If


If UserList(userindex).Stats.GLD < OroEntrada Then Exit Sub


    If UserList(userindex).flags.EnDM Then
        'Call SendData(ToIndex, userindex, 0, "||Ya estás en DeathMatch!" & FONTTYPE_BLANCO)
        'Exit Sub
    NoSave = True
    End If


UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - OroEntrada
Call SendUserORO(userindex)

'ACA DEBERIA GUARDAR EL USUARIO
If UserList(userindex).flags.UserLogged Then
    If NoSave = False Then
        Call SaveUser(userindex, CharPath & UCase$(UserList(userindex).Name) & ".chr")
    End If
Else
    Exit Sub
End If
'LISTO(?)
For i = 1 To UBound(ClaseDM)
    If UCase$(ClaseDM(i).Nombre) = UCase$(Clase) Then
        UserList(userindex).flags.EnDM = True
        UserList(userindex).Stats.MaxHP = ClaseDM(i).Vida
        UserList(userindex).Stats.MaxMAN = ClaseDM(i).Mana
        For x = 1 To NUMCLASES
        If UCase$(Clase) = UCase$(ListaClases(x)) Then
        Exit For
        End If
        DoEvents
        Next x
        UserList(userindex).Clase = x
        
        For x = 1 To NUMRAZAS
        If UCase$(ClaseDM(i).Raza) = UCase$(ListaRazas(x)) Then
        Exit For
        End If
        DoEvents
        Next x
        UserList(userindex).Raza = x

        For x = 1 To NUMSKILLS
        UserList(userindex).Stats.UserSkills(x) = 100
        DoEvents
        Next x
        
        For x = 1 To NUMATRIBUTOS
        UserList(userindex).Stats.UserAtributos(x) = 19
        UserList(userindex).Stats.UserAtributosBackUP(x) = 19
        DoEvents
        Next x
        
        For x = 1 To MAXUSERHECHIZOS
        UserList(userindex).Stats.UserHechizos(x) = ClaseDM(i).Hechizos(x)
        DoEvents
        Next x
        
        UserList(userindex).Genero = HOMBRE
        
        Call UpdateUserHechizos(True, userindex, 0)
        
        UserList(userindex).Stats.ELV = 45
        UserList(userindex).Stats.ELU = 9999999
        UserList(userindex).Stats.Exp = 0
        
        
        UserList(userindex).Faccion.Bando = Neutral
    
        UserList(userindex).Recompensas(1) = ClaseDM(i).Recom(1)
        UserList(userindex).Recompensas(2) = ClaseDM(i).Recom(2)
        UserList(userindex).Recompensas(3) = ClaseDM(i).Recom(3)
        
        UserList(userindex).Stats.MaxHIT = ClaseDM(i).MinG
        UserList(userindex).Stats.MinHIT = ClaseDM(i).MaxG
    
        '
        Call CalcularSta(userindex)
        
        UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
        UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMAN

        Call DarEquipo(userindex, ClaseDM(i).NClase)
        Call WarpUserChar(userindex, MapaDM, 75, 75, True)
        Call SendUserStatsBox(userindex)
        Call SendData(ToIndex, userindex, 0, "||" & DM_MBienvenida & OroEntrada & " monedas de oro." & FONTTYPE_BLANCO)
        Exit For
    End If
DoEvents
Next i




End Sub



Sub ReincorporarDM(userindex As Integer)
If UserList(userindex).Stats.GLD < OroMuerte Then CloseSocket (userindex): Exit Sub
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - OroMuerte
Call SendUserORO(userindex)
Call RevivirUsuarioNPC(userindex)
Call WarpUserChar(userindex, MapaEQ, 50, 50, True)
End Sub


Sub PagarDM(userindex As Integer)
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + OroKill
Call SendUserORO(userindex)
End Sub


Sub DarEquipo(userindex As Integer, ClaseAc As Integer)
If UserList(userindex).flags.EnDM = False Then Exit Sub
Dim i As Integer

Call LimpiarInventario(userindex)

If UserList(userindex).Char.loops = LoopAdEternum Then
    UserList(userindex).Char.FX = 0
    UserList(userindex).Char.loops = 0
End If

    UserList(userindex).Char.Body = 0
    UserList(userindex).Char.ShieldAnim = NingunEscudo
    UserList(userindex).Char.WeaponAnim = NingunArma
    UserList(userindex).Char.CascoAnim = NingunCasco

Call UpdateUserInv(True, userindex, 1)
For i = 1 To ClaseDM(ClaseAc).ObjC
    UserList(userindex).Invent.Object(i).OBJIndex = ClaseDM(ClaseAc).ObjDM(i).Num
    UserList(userindex).Invent.Object(i).Amount = ClaseDM(ClaseAc).ObjDM(i).Cant
    UserList(userindex).Invent.Object(i).Equipped = ClaseDM(ClaseAc).ObjDM(i).Equipado

    Select Case ObjData(UserList(userindex).Invent.Object(i).OBJIndex).ObjType
        Case OBJTYPE_ARMOUR
            Select Case ObjData(UserList(userindex).Invent.Object(i).OBJIndex).SubTipo
                Case OBJTYPE_ARMADURA
                    UserList(userindex).Invent.ArmourEqpSlot = i
                Case OBJTYPE_CASCO
                    UserList(userindex).Invent.CascoEqpSlot = i
                Case OBJTYPE_ESCUDO
                    UserList(userindex).Invent.EscudoEqpSlot = i
            End Select
        Case OBJTYPE_WEAPON
            UserList(userindex).Invent.WeaponEqpSlot = i
        Case OBJTYPE_FLECHAS
            UserList(userindex).Invent.MunicionEqpSlot = i
        Case Else
    End Select
DoEvents
Next i
UserList(userindex).Invent.HerramientaEqpslot = 0
UserList(userindex).Invent.BarcoSlot = 0
  
  
  
  
  Dim MiObj As Obj
  MiObj.Amount = 1000
  MiObj.OBJIndex = DMComida
  If Not MeterItemEnInventario(userindex, MiObj) Then
  End If
  MiObj.Amount = 1000
  MiObj.OBJIndex = DMBebida
  If Not MeterItemEnInventario(userindex, MiObj) Then
  End If
  
Call VerObjetosEquipados(userindex)
'Call ChangeUserBody(ToMap, 0, UserList(userindex).POS.Map, userindex, ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje)



Call UpdateUserInv(True, userindex, 1)
Call SendUserStatsBox(userindex)
End Sub



Sub LimpiarInventario(userindex As Integer)
Dim j As Byte

For j = 1 To MAX_INVENTORY_SLOTS
        UserList(userindex).Invent.Object(j).OBJIndex = 0
        UserList(userindex).Invent.Object(j).Amount = 0
        UserList(userindex).Invent.Object(j).Equipped = 0
Next

UserList(userindex).Invent.NroItems = 0

UserList(userindex).Invent.ArmourEqpObjIndex = 0
UserList(userindex).Invent.ArmourEqpSlot = 0

UserList(userindex).Invent.WeaponEqpObjIndex = 0
UserList(userindex).Invent.WeaponEqpSlot = 0

UserList(userindex).Invent.CascoEqpObjIndex = 0
UserList(userindex).Invent.CascoEqpSlot = 0

UserList(userindex).Invent.EscudoEqpObjIndex = 0
UserList(userindex).Invent.EscudoEqpSlot = 0

UserList(userindex).Invent.HerramientaEqpObjIndex = 0
UserList(userindex).Invent.HerramientaEqpslot = 0

UserList(userindex).Invent.MunicionEqpObjIndex = 0
UserList(userindex).Invent.MunicionEqpSlot = 0

UserList(userindex).Invent.BarcoObjIndex = 0
UserList(userindex).Invent.BarcoSlot = 0

End Sub

