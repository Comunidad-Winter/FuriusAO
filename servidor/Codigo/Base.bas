Attribute VB_Name = "Base"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Public Con As ADODB.Connection

'Public Function NoSirve(Texto As String)
'Dim Valorx As Integer
'For Valorx = 1 To Len(Texto)
'If Mid$(Texto, Valorx, 1) = ";" Or Mid$(Texto, Valorx, 1) = "*" Or Mid$(Texto, Valorx, 1) = "/" Or Mid$(Texto, Valorx, 1) = "=" Or Mid$(Texto, Valorx, 1) = "-" Or Mid$(Texto, Valorx, 1) = "+" Then Sirve = True
'Next Valorx
'End Function



Public Sub CargarDB()
On Error GoTo errhandler
'COMPILAR BETA
If MySql = 0 Then Exit Sub

Set Con = New ADODB.Connection
'Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=mysql.server2.servilinkweb.com.ar;" & " DATABASE=furiusao;" & "UID=furiusao;PWD=leo159753; OPTION=3"
'Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=furiusao;" & "UID=root;PWD=leo159753; OPTION=3"
'Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=furiusao.latencia.com.ar;" & " DATABASE=furiusao;" & "UID=furiusao;PWD=159753; OPTION=3"
'Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=furiusao;" & "UID=root;PWD=igames; OPTION=3"
'Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=furiusao;" & "UID=root;PWD=root; OPTION=3"
Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & "SERVER=localhost;" & " DATABASE=furiusao;" & "UID=root;PWD=igames; OPTION=3"
'

Con.CursorLocation = adUseClient
Con.Open

Exit Sub

errhandler:
    Call LogErrorUrgente("Error en CargarDB: " & Err.Description & " String: " & Con.ConnectionString)
   End

End Sub
Public Function ChangePos(UserName As String) As Boolean
Dim IndexPJ As Long
Dim str As String
If MySql = 0 Then Exit Function
Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(UserName) & "'")
If RS.BOF Or RS.EOF Then Exit Function

IndexPJ = RS!IndexPJ

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `cinit` WHERE IndexPJ=" & IndexPJ)
If RS.BOF Or RS.EOF Then Exit Function

'if nosirve(UserList(IndexPJ).Email) Then Exit Function
'if nosirve(UserList(IndexPJ).Desc) Then Exit Function
str = "UPDATE `cinit` SET"
str = str & " IndexPJ=" & IndexPJ
str = str & ",Email='" & RS!Email & "'"
str = str & ",Genero=" & RS!Genero
str = str & ",Raza=" & RS!Raza
str = str & ",Hogar=" & RS!Hogar
str = str & ",Clase=" & RS!Clase
str = str & ",Codigo='" & RS!codigo & "'"
str = str & ",Descripcion='" & RS!Descripcion & "'"
str = str & ",Head=" & RS!Head
str = str & ",LastIP='" & RS!LastIP & "'"
str = str & ",Mapa=" & ULLATHORPE.Map
str = str & ",X=" & ULLATHORPE.x
str = str & ",Y=" & ULLATHORPE.Y
str = str & " WHERE IndexPJ=" & IndexPJ

Call Con.Execute(str)

Set RS = Nothing

End Function
Public Function ChangeBan(ByVal Name As String, ByVal Baneado As Byte) As Boolean
Dim Orden As String
If MySql Then
Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

Orden = "UPDATE `cflags` SET"
Orden = Orden & " IndexPJ=" & RS!IndexPJ
Orden = Orden & ",Nombre='" & UCase$(Name) & "'"
Orden = Orden & ",Ban=" & Baneado
Orden = Orden & " WHERE IndexPJ=" & RS!IndexPJ

Call Con.Execute(Orden)

Set RS = Nothing
Else
Dim UserFile As String
UserFile = CharPath & UCase$(Name) & ".CHR"
'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
'If val(UserList(userindex).Clase) = 0 Or val(UserList(userindex).Stats.ELV) = 0 Then
'    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(userindex).Name)
'    Exit Function
'End If
Call WriteVar(UserFile, "FLAGS", "Ban", val(Baneado))
End If
End Function
Public Sub SendCharInfo(ByVal UserName As String, userindex As Integer)
Dim Data As String
Dim IndexPJ As Long


If Not ExistePersonaje(UserName) Then Exit Sub

Data = "CHRINFO" & UserName

If MySql = 1 Then
Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(UserName) & "'")
If RS.BOF Or RS.EOF Then Exit Sub

IndexPJ = RS!IndexPJ

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `cinit` WHERE IndexPJ=" & IndexPJ)

If RS.BOF Or RS.EOF Then Exit Sub

Data = Data & "," & ListaRazas(RS!Raza) & "," & ListaClases(RS!Clase) & "," & GeneroLetras(RS!Genero) & ","

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `cstats` WHERE IndexPJ=" & IndexPJ)

If RS.BOF Or RS.EOF Then Exit Sub

Data = Data & RS!ELV & "," & RS!GLD & "," & RS!Banco & ","

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `cguild` WHERE IndexPJ=" & IndexPJ)

If RS.BOF Or RS.EOF Then Exit Sub

Data = Data & RS!FundoClan & "," & RS!ClanFundado & "," _
            & RS!Solicitudes & "," & RS!SolicitudesRechazadas & "," _
            & RS!VecesFueGuildLeader & "," & RS!ClanesParticipo & ","

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `cfaccion` WHERE IndexPJ=" & IndexPJ)

If RS.BOF Or RS.EOF Then Exit Sub

Data = Data & RS!Bando & "," & RS!matados0 & "," & RS!matados1 & "," & RS!matados2

Set RS = Nothing
Else
Dim UserFile As String
UserFile = CharPath & UCase$(UserName) & ".chr"
    'With UserList(NameIndex(UserName))
        Data = Data & "," & ListaRazas(GetVar(UserFile, "INIT", "Raza")) & "," & ListaClases(GetVar(UserFile, "INIT", "Clase")) & "," & GeneroLetras(GetVar(UserFile, "INIT", "Genero")) & ","
        Data = Data & val(GetVar(UserFile, "STATS", "ELV")) & "," & val(GetVar(UserFile, "STATS", "GLD")) & "," & val(GetVar(UserFile, "STATS", "BANCO")) & ","
        Data = Data & val(GetVar(UserFile, UCase$(Actual), "FundoClan")) & "," & GetVar(UserFile, UCase$(Actual), "ClanFundado") & "," _
        & val(GetVar(UserFile, UCase$(Actual), "Solicitudes")) & "," & val(GetVar(UserFile, UCase$(Actual), "SolicitudesRechazadas")) & "," _
        & val(GetVar(UserFile, UCase$(Actual), "VecesFueGuildLeader")) & "," & val(GetVar(UserFile, UCase$(Actual), "ClanesParticipo")) & ","
        Data = Data & val(UserList(userindex).Faccion.Bando) & "," & val(GetVar(UserFile, Actual, "matados0")) & "," & val(GetVar(UserFile, Actual, "matados1")) & val(GetVar(UserFile, Actual, "matados2"))
    'End With
    
    
End If

Call SendData(ToIndex, userindex, 0, Data)

End Sub
Public Sub CerrarDB()
On Error GoTo ErrHandle
If MySql Then
Con.Close
Set Con = Nothing
End If
Exit Sub



ErrHandle:
    Call LogErrorUrgente("Ha surgido un error al cerrar la base de datos MySQL")
    End
    
End Sub
Public Sub SaveUserSQL(userindex As Integer)
On Local Error GoTo ErrHandle
Dim RS As ADODB.Recordset
Dim mUser As user
Dim i As Byte
Dim str As String

mUser = UserList(userindex)

If Len(mUser.Name) = 0 Then Exit Sub
'''if nosirve(mUser.Name) Then Exit Sub

Set RS = New ADODB.Recordset

Set RS = Con.Execute("SELECT * FROM `cflags` WHERE IndexPJ=" & UserList(userindex).IndexPJ)

If RS.BOF Or RS.EOF Then
    Con.Execute ("INSERT INTO `cflags` (NOMBRE) VALUES ('" & UCase$(mUser.Name) & "')")
    Set RS = Nothing
    Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(mUser.Name) & "'")
    UserList(userindex).IndexPJ = RS!IndexPJ
End If

Set RS = Nothing
Dim Pena As Integer

Set RS = Con.Execute("SELECT * FROM `cflags` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
str = "UPDATE `cflags` SET"
str = str & " IndexPJ=" & UserList(userindex).IndexPJ
str = str & ",Nombre='" & UCase$(mUser.Name) & "'"
str = str & ",Ban=" & mUser.flags.Ban
str = str & ",Navegando=" & mUser.flags.Navegando
str = str & ",Envenenado=" & mUser.flags.Envenenado
str = str & ",Silenciado=" & mUser.flags.Silenciado
Pena = CalcularTiempoCarcel(userindex)
str = str & ",Pena=" & Pena
str = str & ",Password='" & mUser.Password & "'"
str = str & ",DenunciasCheat=" & mUser.flags.Denuncias
str = str & ",DenunciasInsulto=" & mUser.flags.DenunciasInsultos
str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `cfaccion` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `cfaccion` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
Set RS = Nothing

str = "UPDATE `cfaccion` SET"

str = str & " IndexPJ=" & UserList(userindex).IndexPJ
str = str & ",Bando=" & mUser.Faccion.Bando
str = str & ",BandoOriginal=" & mUser.Faccion.BandoOriginal
str = str & ",Matados0=" & mUser.Faccion.Matados(0)
str = str & ",Matados1=" & mUser.Faccion.Matados(1)
str = str & ",Matados2=" & mUser.Faccion.Matados(2)
str = str & ",Jerarquia=" & mUser.Faccion.Jerarquia
str = str & ",Ataco1=" & Buleano(mUser.Faccion.Ataco(1) = 1)
str = str & ",Ataco2=" & Buleano(mUser.Faccion.Ataco(2) = 1)
str = str & ",Quests=" & mUser.Faccion.Quests
str = str & ",Torneos=" & mUser.Faccion.Torneos
str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `cguild` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `cguild` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
Set RS = Nothing

str = "UPDATE `cguild` SET"
''if nosirve(mUser.GuildInfo.GuildName) Then Exit Sub
str = str & " IndexPJ=" & UserList(userindex).IndexPJ
str = str & ",Echadas=" & mUser.GuildInfo.echadas
str = str & ",SolicitudesRechazadas=" & mUser.GuildInfo.SolicitudesRechazadas
str = str & ",Guildname='" & mUser.GuildInfo.GuildName & "'"
str = str & ",ClanesParticipo=" & mUser.GuildInfo.ClanesParticipo
str = str & ",Guildpts=" & mUser.GuildInfo.GuildPoints
str = str & ",EsGuildLeader=" & mUser.GuildInfo.EsGuildLeader
str = str & ",Solicitudes=" & mUser.GuildInfo.Solicitudes
str = str & ",VecesFueGuildLeader=" & mUser.GuildInfo.VecesFueGuildLeader
str = str & ",YaVoto=" & mUser.GuildInfo.YaVoto
str = str & ",FundoClan=" & mUser.GuildInfo.FundoClan
str = str & ",ClanFundado='" & mUser.GuildInfo.ClanFundado & "'"
str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `catrib` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `catrib` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
Set RS = Nothing

str = "UPDATE `catrib` SET"
str = str & " IndexPJ=" & UserList(userindex).IndexPJ
For i = 1 To NUMATRIBUTOS
    str = str & ",AT" & i & "=" & mUser.Stats.UserAtributosBackUP(i)
Next i
str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `cskills` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `cskills` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
Set RS = Nothing

str = "UPDATE `cskills` SET"
str = str & " IndexPJ=" & UserList(userindex).IndexPJ

For i = 1 To NUMSKILLS
    str = str & ",SK" & i & "=" & mUser.Stats.UserSkills(i)
Next i

str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `cinit` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `cinit` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
Set RS = Nothing
''if nosirve(mUser.Desc) Then Exit Sub
str = "UPDATE `cinit` SET"
str = str & " IndexPJ=" & UserList(userindex).IndexPJ
str = str & ",Email='" & mUser.Email & "'"
str = str & ",Genero=" & mUser.Genero
str = str & ",Raza=" & mUser.Raza
str = str & ",Hogar=" & mUser.Hogar
str = str & ",Clase=" & mUser.Clase
str = str & ",Codigo='" & mUser.codigo & "'"
str = str & ",Descripcion='" & mUser.Desc & "'"
str = str & ",Head=" & mUser.OrigChar.Head
str = str & ",LastIP='" & mUser.ip & "'"
str = str & ",Mapa=" & mUser.POS.Map
str = str & ",X=" & mUser.POS.x
str = str & ",Y=" & mUser.POS.Y
str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
Call Con.Execute(str)


Set RS = Con.Execute("SELECT * FROM `cstats` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `cstats` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
Set RS = Nothing
 
str = "UPDATE `cstats` SET"
str = str & " IndexPJ=" & UserList(userindex).IndexPJ
str = str & ",GLD=" & mUser.Stats.GLD
str = str & ",BANCO=" & mUser.Stats.Banco
str = str & ",MaxHP=" & mUser.Stats.MaxHP
str = str & ",MinHP=" & mUser.Stats.MinHP
str = str & ",MaxMAN=" & mUser.Stats.MaxMAN
str = str & ",MinMAN=" & mUser.Stats.MinMAN
str = str & ",MinSTA=" & mUser.Stats.MinSta
str = str & ",MaxHIT=" & mUser.Stats.MaxHIT
str = str & ",MinHIT=" & mUser.Stats.MinHIT
str = str & ",MaxAGU=" & mUser.Stats.MaxAGU
str = str & ",MinAGU=" & mUser.Stats.MinAGU
str = str & ",MaxHAM=" & mUser.Stats.MaxHam
str = str & ",MinHAM=" & mUser.Stats.MinHam
str = str & ",SkillPtsLibres=" & mUser.Stats.SkillPts
str = str & ",VecesMurioUsuario=" & mUser.Stats.VecesMurioUsuario
str = str & ",EXP=" & mUser.Stats.Exp
str = str & ",ELV=" & mUser.Stats.ELV
str = str & ",NpcsMuertes=" & mUser.Stats.NPCsMuertos
For i = 1 To 3
    str = str & ",Recompensa" & i & "=" & mUser.Recompensas(i)
Next i
str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
 Call Con.Execute(str)

 
 Set RS = Con.Execute("SELECT * FROM `cbanco` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
 If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `cbanco` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
 
 str = "UPDATE `cbanco` SET"
 str = str & " IndexPJ=" & UserList(userindex).IndexPJ
 For i = 1 To MAX_BANCOINVENTORY_SLOTS
     str = str & ",OBJ" & i & "=" & mUser.BancoInvent.Object(i).OBJIndex
     str = str & ",CANT" & i & "=" & mUser.BancoInvent.Object(i).Amount
 Next i
 str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
 Call Con.Execute(str)

 
 Set RS = Con.Execute("SELECT * FROM `chechizos` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
 If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `chechizos` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
 Set RS = Nothing
 
 str = "UPDATE `chechizos` SET"
 str = str & " IndexPJ=" & UserList(userindex).IndexPJ
 For i = 1 To MAXUSERHECHIZOS
     str = str & ",H" & i & "=" & mUser.Stats.UserHechizos(i)
 Next i
 str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
 Call Con.Execute(str)
 
 
 Set RS = Con.Execute("SELECT * FROM `cinvent` WHERE IndexPJ=" & UserList(userindex).IndexPJ)
 If RS.BOF Or RS.EOF Then Call Con.Execute("INSERT INTO `cinvent` (IndexPJ) VALUES (" & UserList(userindex).IndexPJ & ")")
 Set RS = Nothing
 
 str = "UPDATE `cinvent` SET"
 str = str & " IndexPJ=" & UserList(userindex).IndexPJ
 For i = 1 To MAX_INVENTORY_SLOTS
     str = str & ",OBJ" & i & "=" & mUser.Invent.Object(i).OBJIndex
     str = str & ",CANT" & i & "=" & mUser.Invent.Object(i).Amount
 Next i
 str = str & ",CASCOSLOT=" & mUser.Invent.CascoEqpSlot
str = str & ",ARMORSLOT=" & mUser.Invent.ArmourEqpSlot
str = str & ",SHIELDSLOT=" & mUser.Invent.EscudoEqpSlot
str = str & ",WEAPONSLOT=" & mUser.Invent.WeaponEqpSlot
str = str & ",HERRAMIENTASLOT=" & mUser.Invent.HerramientaEqpslot
str = str & ",MUNICIONSLOT=" & mUser.Invent.MunicionEqpSlot
str = str & ",BARCOSLOT=" & mUser.Invent.BarcoSlot
 
 str = str & " WHERE IndexPJ=" & UserList(userindex).IndexPJ
 Call Con.Execute(str)

Call RevisarTops(userindex)

Exit Sub

ErrHandle:
    Call LogErrorUrgente("Error en SaveUserSQL: " & Err.Description & " String: " & Con.ConnectionString)
    Resume Next
End Sub
Function CalcularTiempoCarcel(userindex As Integer) As Integer

If UserList(userindex).flags.Encarcelado = 1 Then CalcularTiempoCarcel = 1 + (UserList(userindex).Counters.TiempoPena - TiempoTranscurrido(UserList(userindex).Counters.Pena)) \ 60

End Function
Function LoadUserSQL(userindex As Integer, ByVal Name As String) As Boolean
On Error GoTo errhandler
Dim i As Integer

With UserList(userindex)
    Dim RS As New ADODB.Recordset
    Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If

    .IndexPJ = RS!IndexPJ
    Set RS = Nothing
    
    Set RS = Con.Execute("SELECT * FROM `cflags` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    'if nosirve(UserList(userindex).Password) Then Exit Function
    
    
    .flags.Ban = RS!Ban
    .flags.Navegando = RS!Navegando
    .flags.Envenenado = RS!Envenenado
    .Counters.TiempoPena = RS!Pena * 60
    .Password = RS!Password
    .flags.Denuncias = RS!DenunciasCheat
    .flags.DenunciasInsultos = RS!DenunciasInsulto
    .flags.Silenciado = RS!Silenciado

    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `cfaccion` WHERE IndexPJ=" & .IndexPJ)
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    .Faccion.Bando = RS!Bando
    .Faccion.BandoOriginal = RS!BandoOriginal
    .Faccion.Matados(0) = RS!matados0
    .Faccion.Matados(1) = RS!matados1
    .Faccion.Matados(2) = RS!matados2
    .Faccion.Jerarquia = RS!Jerarquia
    .Faccion.Ataco(1) = RS!Ataco1
    .Faccion.Ataco(2) = RS!Ataco2
    .Faccion.Quests = RS!Quests
    .Faccion.Torneos = RS!Torneos
    Set RS = Nothing

    If Not ModoQuest And UserList(userindex).Faccion.Bando <> Neutral And UserList(userindex).Faccion.Bando <> UserList(userindex).Faccion.BandoOriginal Then UserList(userindex).Faccion.Bando = Neutral

    Set RS = Con.Execute("SELECT * FROM `cguild` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    'if nosirve(UserList(userindex).GuildInfo.GuildName) Then Exit Function
    .GuildInfo.EsGuildLeader = RS!EsGuildLeader
    .GuildInfo.echadas = RS!echadas
    .GuildInfo.Solicitudes = RS!Solicitudes
    .GuildInfo.SolicitudesRechazadas = RS!SolicitudesRechazadas
    .GuildInfo.VecesFueGuildLeader = RS!VecesFueGuildLeader
    .GuildInfo.YaVoto = RS!YaVoto
    .GuildInfo.FundoClan = RS!FundoClan
    .GuildInfo.GuildName = RS!GuildName
    .GuildInfo.ClanFundado = RS!ClanFundado
    .GuildInfo.ClanesParticipo = RS!ClanesParticipo
    .GuildInfo.GuildPoints = RS!GuildPts
    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `catrib` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    For i = 1 To NUMATRIBUTOS
        .Stats.UserAtributos(i) = RS.Fields("AT" & i)
        .Stats.UserAtributosBackUP(i) = .Stats.UserAtributos(i)
    Next i
    
    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `cskills` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To NUMSKILLS
        .Stats.UserSkills(i) = RS.Fields("SK" & i)
    Next i
    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `cbanco` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        .BancoInvent.Object(i).OBJIndex = RS.Fields("OBJ" & i)
        .BancoInvent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    Set RS = Nothing
    
    
    Set RS = Con.Execute("SELECT * FROM `cinvent` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAX_INVENTORY_SLOTS
        .Invent.Object(i).OBJIndex = RS.Fields("OBJ" & i)
        .Invent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    .Invent.CascoEqpSlot = RS!CASCOSLOT
    .Invent.ArmourEqpSlot = RS!ARMORSLOT
    .Invent.EscudoEqpSlot = RS!SHIELDSLOT
    .Invent.WeaponEqpSlot = RS!WEAPONSLOT
    .Invent.HerramientaEqpslot = RS!HERRAMIENTASLOT
    .Invent.MunicionEqpSlot = RS!MUNICIONSLOT
    .Invent.BarcoSlot = RS!BarcoSlot
    Set RS = Nothing

    
    Set RS = Con.Execute("SELECT * FROM `chechizos` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAXUSERHECHIZOS
        .Stats.UserHechizos(i) = RS.Fields("H" & i)
    Next i
    Set RS = Nothing
    
    Set RS = Con.Execute("SELECT * FROM `cstats` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    .Stats.GLD = RS!GLD
    .Stats.Banco = RS!Banco
    .Stats.MaxHP = RS!MaxHP
    .Stats.MinHP = RS!MinHP
    .Stats.MinSta = RS!MinSta
    .Stats.MaxMAN = RS!MaxMAN
    .Stats.MinMAN = RS!MinMAN
    .Stats.MaxHIT = RS!MaxHIT
    .Stats.MinHIT = RS!MinHIT
    .Stats.MinAGU = RS!MinAGU
    .Stats.MinHam = RS!MinHam
    .Stats.SkillPts = RS!SkillPtsLibres
    .Stats.VecesMurioUsuario = RS!VecesMurioUsuario
    .Stats.Exp = RS!Exp
    .Stats.ELV = RS!ELV
    .Stats.ELU = ELUs(.Stats.ELV)
    .Stats.NPCsMuertos = RS!NpcsMuertes

    For i = 1 To 3
        .Recompensas(i) = RS.Fields("Recompensa" & i)
    Next
    
    Set RS = Nothing
    
    If .Stats.MinAGU < 1 Then .flags.Sed = 1
    If .Stats.MinHam < 1 Then .flags.Hambre = 1
    If .Stats.MinHP < 1 Then .flags.Muerto = 1
        
    'if nosirve(UserList(userindex).Desc) Then Exit Function
    Set RS = Con.Execute("SELECT * FROM `cinit` WHERE IndexPJ=" & .IndexPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    .Email = RS!Email
    .Genero = RS!Genero
    .Raza = RS!Raza
    .Hogar = RS!Hogar
    .Clase = RS!Clase
    .codigo = RS!codigo
    .Desc = RS!Descripcion
    .OrigChar.Head = RS!Head
    .POS.Map = RS!mapa
    .POS.x = RS!x
    .POS.Y = RS!Y

    If .flags.Muerto = 0 Then
        .Char = .OrigChar
        Call VerObjetosEquipados(userindex)
    Else
        .Char.Body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
    End If
    
    .Char.Heading = 3
    
    Set RS = Nothing
    
    LoadUserSQL = True


    If Len(.Desc) >= 80 Then .Desc = Left$(.Desc, 80)

    If .Counters.TiempoPena > 0 Then
        .flags.Encarcelado = 1
        .Counters.Pena = Timer
    End If
    
    .Stats.MaxAGU = 100
    .Stats.MaxHam = 100
    Call CalcularSta(userindex)

End With

Exit Function

errhandler:
    Call LogError("Error en LoadUserSQL. N:" & Name & " - " & Err.Number & "-" & Err.Description)
    Set RS = Nothing
    
End Function
Function SumarDenuncia(ByVal Name As String, Tipo As Byte) As Integer
Dim RS As New ADODB.Recordset
On Error GoTo Error
Dim str As String, Den As Integer

Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

str = "UPDATE `cflags` SET"
str = str & " IndexPJ=" & RS!IndexPJ
str = str & ",Nombre='" & RS!Nombre & "'"
str = str & ",Ban=" & RS!Ban
str = str & ",Navegando=" & RS!Navegando
str = str & ",Envenenado=" & RS!Envenenado
str = str & ",Pena=" & RS!Pena
str = str & ",Password='" & RS!Password & "'"
'if nosirve(SumarDenuncia & DenunciasInsulto) Then Exit Function
If Tipo = 1 Then
    Den = RS!DenunciasCheat
    SumarDenuncia = Den + 1
    str = str & ",DenunciasCheat=" & SumarDenuncia
    str = str & ",DenunciasInsulto=" & RS!DenunciasInsulto
Else
    Den = RS!DenunciasInsulto
    SumarDenuncia = Den + 1
    str = str & ",DenunciasCheat=" & RS!DenunciasCheat
    str = str & ",DenunciasInsulto=" & SumarDenuncia
End If

str = str & " WHERE IndexPJ=" & RS!IndexPJ
Call Con.Execute(str)

Set RS = Nothing
Exit Function
Error:
    Call LogError("Error en SumarDenuncia: " & Err.Description & " " & Name & " " & Tipo)
    
End Function
Function ComprobarPassword(ByVal Name As String, Password As String, Optional Maestro As Boolean) As Byte
Dim Pass As String
If MySql Then
Dim RS As New ADODB.Recordset
Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")
If RS.BOF Or RS.EOF Then Exit Function

Pass = RS!Password
Set RS = Nothing
Else
Pass = GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "Password")
'If GetVar(CharPath & UCase$(Name) & ".chr", "FLAGS", "PIN") = Password Then
'    ComprobarPassword = True
'    Exit Function
'End If
End If
If Len(Pass) = 0 Then Exit Function
ComprobarPassword = (Password = Pass)




End Function
Public Function BANCheck(ByVal Name As String) As Boolean
Dim RS As New ADODB.Recordset
Dim Baneado As Byte


If MySql Then
Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

Baneado = RS!Ban
BANCheck = (Baneado = 1)

Set RS = Nothing
Else

BANCheck = val(GetVar(CharPath & Name & ".CHR", "FLAGS", "Ban"))
'BANCheck = (BANCheck = 1)
'end if
End If
End Function
Public Function IndexPJ(ByVal Name As String) As Integer
Dim RS As New ADODB.Recordset
Dim Baneado As Byte

Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

IndexPJ = RS!IndexPJ

Set RS = Nothing

End Function
Function ExistePersonaje(Name As String) As Boolean
Dim RS As New ADODB.Recordset
If Len(Name) = 0 Then Exit Function
If MySql Then
 
Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

Set RS = Nothing

ExistePersonaje = True
Else
If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then ExistePersonaje = True

End If


End Function
Function AgregarAClan(ByVal Name As String, ByVal Clan As String) As Boolean
Dim RS As New ADODB.Recordset
Dim IndexPJ As Long
Dim str As String
If MySql Then
Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Function

IndexPJ = RS!IndexPJ

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `cguild` WHERE IndexPJ=" & IndexPJ)
If RS.BOF Or RS.EOF Then Exit Function
'if nosirve(Clan) Then Exit Function
If Len(RS!GuildName) = 0 Then
    str = "UPDATE `cguild` SET"
    str = str & " IndexPJ=" & IndexPJ
    str = str & ",Echadas=" & RS!echadas
    str = str & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas
    str = str & ",Guildname='" & Clan & "'"
    str = str & ",ClanesParticipo=" & RS!ClanesParticipo + 1
    str = str & ",Guildpts=" & RS!GuildPts + 25
    str = str & " WHERE IndexPJ=" & IndexPJ
    Call Con.Execute(str)
    AgregarAClan = True
End If

Set RS = Nothing
Else
Dim UserfileX As String
UserfileX = CharPath & Name & ".CHR"
'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If Len(GetVar(UserfileX, "GUILD", "GuildName")) = 0 Then
Call WriteVar(UserfileX, "GUILD", "GuildName", Clan)
Call WriteVar(UserfileX, "GUILD", "ClanesParticipo", val(GetVar(UserfileX, "GUILD", "ClanesParticipo")) + 1)
Call WriteVar(UserfileX, "GUILD", "GuildPts", val(GetVar(UserfileX, "GUILD", "GuildPts")) + 25)
 AgregarAClan = True
End If
End If
End Function
Sub RechazarSolicitud(ByVal Name As String)
If MySql Then
Dim RS As New ADODB.Recordset
Dim IndexPJ As Long
Dim Orden As String

Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Sub

IndexPJ = RS!IndexPJ

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `cguild` WHERE IndexPJ=" & IndexPJ)
If RS.BOF Or RS.EOF Then Exit Sub

Orden = "UPDATE `cguild` SET"
Orden = Orden & " IndexPJ=" & IndexPJ
Orden = Orden & ",Echadas=" & RS!echadas
Orden = Orden & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas + 1
Orden = Orden & " WHERE IndexPJ=" & IndexPJ
Call Con.Execute(Orden)

Set RS = Nothing
Else
'UserList(NameIndex(Name)).GuildInfo.SolicitudesRechazadas = UserList(NameIndex(Name)).GuildInfo.SolicitudesRechazadas + 1
'PARA QUE GUARDARLO :s AL RE PEDO
Dim UserFile As String
UserFile = CharPath & Name & ".CHR"
Call WriteVar(UserFile, "GUILD", "SolicitudesRechazadas", val(GetVar(UserFile, "GUILD", "SolicitudesRechazadas")) + 1)
End If


End Sub
Sub EcharDeClan(ByVal Name As String)
Dim RS As New ADODB.Recordset
Dim IndexPJ As Long
Dim str As String
Dim Echa As Integer
If MySql Then
Set RS = Con.Execute("SELECT * FROM `cflags` WHERE Nombre='" & UCase$(Name) & "'")

If RS.BOF Or RS.EOF Then Exit Sub

IndexPJ = RS!IndexPJ

Set RS = Nothing

Set RS = Con.Execute("SELECT * FROM `cguild` WHERE IndexPJ=" & IndexPJ)
If RS.BOF Or RS.EOF Then Exit Sub

str = "UPDATE `cguild` SET"
str = str & " IndexPJ=" & IndexPJ
Echa = RS!echadas
Echa = Echa + 1
str = str & ",Echadas=" & Echa
str = str & ",SolicitudesRechazadas=" & RS!SolicitudesRechazadas
str = str & ",Guildname=''"
str = str & " WHERE IndexPJ=" & IndexPJ

Call Con.Execute(str)

Set RS = Nothing
Else
Dim UserFiL As String
UserFiL = CharPath & Name & ".CHR"
'UserList(userindex).GuildInfo.echadas = UserList(userindex).GuildInfo.echadas + 1
Call WriteVar(UserFiL, "GUILD", "Echadas", val(GetVar(UserFiL, "GUILD", "Echadas")) + 1)
Call WriteVar(UserFiL, "GUILD", "Guildname", "")

End If


End Sub
