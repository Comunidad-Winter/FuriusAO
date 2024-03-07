Attribute VB_Name = "Seguridad"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal wIndx As Long) As Long

'ProccExes
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


'Close
'FIXIT: As Any no se admite en Visual Basic .NET. Utilice un tipo específico.              FixIT90210ae-R5608-H1984
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'/Close


Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 260
End Type

'/fin








Public ModoProcesos As Boolean
Public PalabrasP() As String






'CERRAR EXE


Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_CREATE_THREAD = &H2
Public Const PROCESS_VM_OPERATION = &H8
Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_VM_WRITE = &H20
Public Const PROCESS_DUP_HANDLE = &H40
Public Const PROCESS_CREATE_PROCESS = &H80
Public Const PROCESS_SET_QUOTA = &H100
Public Const PROCESS_SET_INFORMATION = &H200
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF


'/FIN




Const GWL_STYLE = (-16)
Const Win_VISIBLE = &H10000000
Const Win_BORDER = &H800000
Const SC_CLOSE = &HF060&
Const WM_SYSCOMMAND = &H112



Public Function BanPC(Mac As String) As Boolean
Dim Cant As Integer
Cant = val(GetVar(App.Path & "/logs/BanPC.txt", "BANS", "Cantidad"))
Dim ix As Integer
For ix = 1 To Cant
    If GetVar(App.Path & "/logs/BanPC.txt", "BANS", "Ban" & ix) = Mac Then
        BanPC = True
        Exit Function
    End If
DoEvents
Next ix
End Function

Public Sub CloseExe(tit As String)
hProcess = OpenProcess(PROCESS_TERMINATE, True, tit)
iResult = TerminateProcess(hProcess, 99)
CloseHandle hProcess
End Sub



Public Sub CloseApp(ByVal Titulo As String, Optional ClassName As String)
Call SendMessage(FindWindow(ClassName, Titulo), WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)
End Sub

Public Sub ProccExes()
Dim hSnapShot As Long
Dim uProceso As PROCESSENTRY32
Dim TotalPr As String
Dim Res As Long
hSnapShot = CreateToolhelpSnapshot(2&, 0&)
If hSnapShot <> 0 Then
uProceso.dwSize = Len(uProceso)
Res = ProcessFirst(hSnapShot, uProceso)
Do While Res
TotalPr = TotalPr & " ," & Left$(uProceso.szExeFile, InStr(uProceso.szExeFile, Chr$(0)) - 1) & ":" & uProceso.th32ProcessID
Res = ProcessNext(hSnapShot, uProceso)
Loop
Call CloseHandle(hSnapShot)
End If

Call SendData(ToAdmins, 0, 0, "||Exes de la pc " & TotalPr & FONTTYPE_CELESTE)
End Sub

Public Sub EnumerarV()
Dim IsTask As Long, hwCurr As Long, intLen As Long, strTitle As String
Dim StringTotal As String
IsTask = Win_VISIBLE Or Win_BORDER
hwCurr = GetWindow(frmMain.hwnd, 0)
    Do While hwCurr
        If hwCurr <> frmMain.hwnd And (GetWindowLong(hwCurr, GWL_STYLE) And IsTask) = IsTask Then
            intLen = GetWindowTextLength(hwCurr) + 1
            strTitle = Space$(intLen)
            intLen = GetWindowText(hwCurr, strTitle, intLen)
           ' If intLen > 0 Then
'FIXIT: Reexmplazar la función 'Trim' con la función 'Trim$'.                               FixIT90210ae-R9757-R1B8ZE
                StringTotal = StringTotal & ", " & Left$(Trim(strTitle), 10)
            'End If
        End If
        hwCurr = GetWindow(hwCurr, 2)
    Loop
Call SendData(ToAdmins, 0, 0, "||Procesos de la PC " & StringTotal & FONTTYPE_CELESTE)

    
'Form1.Enivar "Terminado"
End Sub


Public Sub CheckearDevolucion()
Dim Check As Integer
For Check = 1 To LastUser
    If UserList(Check).flags.Devolvio = False Then
    If UserList(Check).flags.UserLogged = True Then
            Call SendData(ToAdmins, 0, 0, "||" & UserList(Check).Name & ": Está con modificado o esta loggeado" & FONTTYPE_INFO)
            Call CloseSocket(Check)
    End If
    End If
DoEvents
Next Check
End Sub

Public Sub CheckearDevoluciones()
Dim Check As Integer
For Check = 1 To LastUser
    If UserList(Check).flags.Devolvio = False Then
    If UserList(Check).flags.UserLogged = True Then
            Call SendData(ToAdmins, 0, 0, "||" & UserList(Check).Name & ": Está con modificado o log." & FONTTYPE_INFO)
    End If
    End If
DoEvents
Next Check
End Sub




Public Sub CheckearDevolucion1(ByVal userindex As Integer)
If UserList(userindex).flags.Devolvio = False Then
    Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & ": Está con un cliente no oficial" & FONTTYPE_INFO)
    Else
    Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & ": Está con un cliente oficial" & FONTTYPE_INFO)
End If
End Sub

Public Function EncriptarString(ByVal Cadena As String) As String
'Dim X As Integer
'For X = 1 To Len(Cadena)
'EncriptarString = EncriptarString & Chr(Asc(Mid$(Cadena, X, 1)) + X)
'Next X
EncriptarString = Cadena
End Function


Public Function DesencriptarFPS(ByVal Cadena As String) As String
Dim x As Integer
For x = 1 To Len(Cadena)
DesencriptarFPS = DesencriptarFPS & Chr(Asc(Mid$(Cadena, x, 1)) - x - 5)
Next x
End Function


Public Sub LimpiarMapas()
'FIXIT: Declare 'MapaActual' and 'Xfd' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim MapaActual, Xfd, Yfd As Integer
MapaActual = 1
'Leito ta el paeod esto solo molesta xD
'Call SendData(ToAll, 0, 0, "||****Limpiando todos los mapas****" & FONTTYPE_VENENO)
For MapaActual = 1 To NumMaps
    For Yfd = YMinMapSize To YMaxMapSize
        For Xfd = XMinMapSize To XMaxMapSize
                If MapData(MapaActual, Xfd, Yfd).OBJInfo.OBJIndex > 0 Then
                If Not ItemEsDeMapa(val(MapaActual), val(Xfd), val(Yfd)) Or MapData(MapaActual, Xfd, Yfd).OBJInfo.OBJIndex = FOGATA Then
                Call EraseObj(ToMap, userindex, MapaActual, 10000, val(MapaActual), val(Xfd), val(Yfd))
                End If
                End If
        Next Xfd
    Next Yfd
Next MapaActual
Call SendData(ToAll, 0, 0, "||¡El mundo fue limpiado!" & FONTTYPE_furius)
End Sub



'LOAD USER^^


Sub SaveFurius(ByVal userindex As Integer, ByVal UserFile As String)
On Error GoTo errhandler
'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If val(UserList(userindex).Stats.ELV) = 0 Then
    Call LogCriticEvent("Save Furius: Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(userindex).Name)
    Exit Sub
End If



If FileExist(UserFile, vbNormal) Then
       If UserList(userindex).flags.Muerto = 1 Then
        OldUserHead = UserList(userindex).Char.Head
        UserList(userindex).Char.Head = val(GetVar(UserFile, "INIT", "Head"))
       End If
End If

Dim LoopC As Integer





'SISTEMA NUEVO DE OBJETOS Y BANCO
Call WriteVar(UserFile, "FLAGS", "SistemaNuevo", val(1))
'/



Call WriteVar(UserFile, "FLAGS", "Pena", val(CalcularTiempoCarcel(userindex)))
Call WriteVar(UserFile, "FLAGS", "PenaM", val(UserList(userindex).Counters.PenaMinar))
'UserList(userindex).Counters.PenaMinar
Call WriteVar(UserFile, "FLAGS", "Password", UserList(userindex).Password)
Call WriteVar(UserFile, "FLAGS", "Silenciado", val(UserList(userindex).flags.Silenciado))

Call WriteVar(UserFile, "FACCIONES", "Bando", val(UserList(userindex).Faccion.Bando))
Call WriteVar(UserFile, "FACCIONES", "BandoOriginal", val(UserList(userindex).Faccion.BandoOriginal))
Call WriteVar(UserFile, "FACCIONES", "Matados0", val(UserList(userindex).Faccion.Matados(0)))
Call WriteVar(UserFile, "FACCIONES", "Matados1", val(UserList(userindex).Faccion.Matados(1)))
Call WriteVar(UserFile, "FACCIONES", "Matados2", val(UserList(userindex).Faccion.Matados(2)))
Call WriteVar(UserFile, "FACCIONES", "Jerarquia", val(UserList(userindex).Faccion.Jerarquia))
Call WriteVar(UserFile, "FACCIONES", "Quests", val(UserList(userindex).Faccion.Quests))
Call WriteVar(UserFile, "FACCIONES", "Torneos", val(UserList(userindex).Faccion.Torneos))
Call WriteVar(UserFile, "FACCIONES", "RGanados", val(UserList(userindex).flags.MatadasenR))
Call WriteVar(UserFile, "FACCIONES", "RPerdidos", val(UserList(userindex).flags.PerdidasenR))


Call WriteVar(UserFile, "GUILD", "GuildName", UserList(userindex).GuildInfo.GuildName)
Call WriteVar(UserFile, "GUILD", "EsGuildLeader", val(UserList(userindex).GuildInfo.EsGuildLeader))
Call WriteVar(UserFile, "GUILD", "FundoClan", val(UserList(userindex).GuildInfo.FundoClan))
Call WriteVar(UserFile, "GUILD", "ClanFundado", UserList(userindex).GuildInfo.ClanFundado)



'¿Fueron modificados los atributos del usuario?
    For LoopC = 1 To NUMATRIBUTOS
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(userindex).Stats.UserAtributosBackUP(LoopC)))
    Next LoopC

Call WriteVar(UserFile, "INIT", "Clase", val(UserList(userindex).Clase))
Call WriteVar(UserFile, "INIT", "Mapa", val(UserList(userindex).POS.Map))
Call WriteVar(UserFile, "INIT", "X", val(UserList(userindex).POS.x))
Call WriteVar(UserFile, "INIT", "Y", val(UserList(userindex).POS.Y))

Call WriteVar(UserFile, "STATS", "GLD", val(UserList(userindex).Stats.GLD))
Call WriteVar(UserFile, "STATS", "BANCO", val(UserList(userindex).Stats.Banco))

Call WriteVar(UserFile, "STATS", "MaxHP", val(UserList(userindex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MaxMAN", val(UserList(userindex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MaxHIT", val(UserList(userindex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", val(UserList(userindex).Stats.MinHIT))
Call WriteVar(UserFile, "STATS", "EXP", val(UserList(userindex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "ELV", val(UserList(userindex).Stats.ELV))
Call WriteVar(UserFile, "STATS", "NpcsMuertes", val(UserList(userindex).Stats.NPCsMuertos))

For LoopC = 1 To 3
Call WriteVar(UserFile, "STATS", "Recompensa" & LoopC, val(UserList(userindex).Recompensas(LoopC)))
Next LoopC

Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    'Call WriteVar(UserFile, "BANCO", "OBJ" & loopd, val(UserList(userindex).BancoInvent.Object(loopd).OBJIndex))
    'Call WriteVar(UserFile, "BANCO", "CANT" & loopd, val(UserList(userindex).BancoInvent.Object(loopd).Amount))
Call WriteVar(UserFile, "BANCO", "OBJ" & loopd, val(UserList(userindex).BancoInvent.Object(loopd).OBJIndex) & "-" & val(UserList(userindex).BancoInvent.Object(loopd).Amount))
Next loopd

Dim cad As String
For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(userindex).Stats.UserHechizos(LoopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
Next LoopC


For LoopC = 1 To MAX_INVENTORY_SLOTS
    'Call WriteVar(UserFile, "Invent", "OBJ" & LoopC, val(UserList(userindex).Invent.Object(LoopC).OBJIndex))
    'Call WriteVar(UserFile, "Invent", "CANT" & LoopC, val(UserList(userindex).Invent.Object(LoopC).Amount))
Call WriteVar(UserFile, "Invent", "OBJ" & LoopC, val(UserList(userindex).Invent.Object(LoopC).OBJIndex) & "-" & val(UserList(userindex).Invent.Object(LoopC).Amount))
Next LoopC


Exit Sub
errhandler:
Call LogError("Error en SaveFurius")

End Sub




Sub SaveUser(ByVal userindex As Integer, ByVal UserFile As String)
On Error GoTo errhandler
'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If UserList(userindex).flags.EnDM = True Then Exit Sub
If val(UserList(userindex).Stats.ELV) = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(userindex).Name)
    Exit Sub
End If



If FileExist(UserFile, vbNormal) Then
       If UserList(userindex).flags.Muerto = 1 Then
        OldUserHead = UserList(userindex).Char.Head
        UserList(userindex).Char.Head = val(GetVar(UserFile, "INIT", "Head"))
       End If
End If

Dim LoopC As Integer

'SISTEMA NUEVO DE OBJETOS Y BANCO
Call WriteVar(UserFile, "FLAGS", "SistemaNuevo", val(1))
'/


Call WriteVar(UserFile, "FLAGS", "Ban", val(UserList(userindex).flags.Ban))
Call WriteVar(UserFile, "FLAGS", "Navegando", val(UserList(userindex).flags.Navegando))
Call WriteVar(UserFile, "FLAGS", "Envenenado", val(UserList(userindex).flags.Envenenado))
Call WriteVar(UserFile, "FLAGS", "Pena", val(CalcularTiempoCarcel(userindex)))
Call WriteVar(UserFile, "FLAGS", "PenaM", val(UserList(userindex).Counters.PenaMinar))

'UserList(userindex).Counters.PenaMinar
Call WriteVar(UserFile, "FLAGS", "DenunciasCheat", val(UserList(userindex).flags.Denuncias))
Call WriteVar(UserFile, "FLAGS", "DenunciasInsulto", val(UserList(userindex).flags.DenunciasInsultos))
Call WriteVar(UserFile, "FLAGS", "Password", UserList(userindex).Password)
Call WriteVar(UserFile, "FLAGS", "PIN", UserList(userindex).PIN)
Call WriteVar(UserFile, "FLAGS", "Silenciado", val(UserList(userindex).flags.Silenciado))


Call WriteVar(UserFile, "FACCIONES", "Bando", val(UserList(userindex).Faccion.Bando))
Call WriteVar(UserFile, "FACCIONES", "BandoOriginal", val(UserList(userindex).Faccion.BandoOriginal))
Call WriteVar(UserFile, "FACCIONES", "Matados0", val(UserList(userindex).Faccion.Matados(0)))
Call WriteVar(UserFile, "FACCIONES", "Matados1", val(UserList(userindex).Faccion.Matados(1)))
Call WriteVar(UserFile, "FACCIONES", "Matados2", val(UserList(userindex).Faccion.Matados(2)))
Call WriteVar(UserFile, "FACCIONES", "Jerarquia", val(UserList(userindex).Faccion.Jerarquia))
Call WriteVar(UserFile, "FACCIONES", "Ataco1", val(UserList(userindex).Faccion.Ataco(1)))
Call WriteVar(UserFile, "FACCIONES", "Ataco2", val(UserList(userindex).Faccion.Ataco(2)))
Call WriteVar(UserFile, "FACCIONES", "Quests", val(UserList(userindex).Faccion.Quests))
Call WriteVar(UserFile, "FACCIONES", "Torneos", val(UserList(userindex).Faccion.Torneos))
Call WriteVar(UserFile, "FACCIONES", "RGanados", val(UserList(userindex).flags.MatadasenR))
Call WriteVar(UserFile, "FACCIONES", "RPerdidos", val(UserList(userindex).flags.PerdidasenR))


Call WriteVar(UserFile, "GUILD", "Echadas", val(UserList(userindex).GuildInfo.echadas))
Call WriteVar(UserFile, "GUILD", "SolicitudesRechazadas", val(UserList(userindex).GuildInfo.SolicitudesRechazadas))
Call WriteVar(UserFile, "GUILD", "GuildName", UserList(userindex).GuildInfo.GuildName)
Call WriteVar(UserFile, "GUILD", "ClanesParticipo", val(UserList(userindex).GuildInfo.ClanesParticipo))
Call WriteVar(UserFile, "GUILD", "GuildPts", val(UserList(userindex).GuildInfo.GuildPoints))
Call WriteVar(UserFile, "GUILD", "EsGuildLeader", val(UserList(userindex).GuildInfo.EsGuildLeader))
Call WriteVar(UserFile, "GUILD", "Solicitudes", val(UserList(userindex).GuildInfo.Solicitudes))
Call WriteVar(UserFile, "GUILD", "VecesFueGuildLeader", val(UserList(userindex).GuildInfo.VecesFueGuildLeader))
Call WriteVar(UserFile, "GUILD", "YaVoto", val(UserList(userindex).GuildInfo.YaVoto))
Call WriteVar(UserFile, "GUILD", "FundoClan", val(UserList(userindex).GuildInfo.FundoClan))
Call WriteVar(UserFile, "GUILD", "ClanFundado", UserList(userindex).GuildInfo.ClanFundado)



'¿Fueron modificados los atributos del usuario?
    For LoopC = 1 To NUMATRIBUTOS
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(userindex).Stats.UserAtributosBackUP(LoopC)))
    Next LoopC

For LoopC = 1 To NUMSKILLS
    Call WriteVar(UserFile, "SKILLS", "SK" & LoopC, val(UserList(userindex).Stats.UserSkills(LoopC)))
Next LoopC


Call WriteVar(UserFile, "INIT", "Email", UserList(userindex).Email)
Call WriteVar(UserFile, "INIT", "Genero", val(UserList(userindex).Genero))
Call WriteVar(UserFile, "INIT", "Raza", val(UserList(userindex).Raza))
Call WriteVar(UserFile, "INIT", "Hogar", val(UserList(userindex).Hogar))
Call WriteVar(UserFile, "INIT", "Clase", val(UserList(userindex).Clase))
Call WriteVar(UserFile, "INIT", "Descripcion", UserList(userindex).Desc)
Call WriteVar(UserFile, "INIT", "Motivo", UserList(userindex).Moti)

Call WriteVar(UserFile, "INIT", "Head", val(UserList(userindex).OrigChar.Head))
Call WriteVar(UserFile, "INIT", "LastIP", UserList(userindex).ip)
Call WriteVar(UserFile, "INIT", "Mapa", val(UserList(userindex).POS.Map))
Call WriteVar(UserFile, "INIT", "X", val(UserList(userindex).POS.x))
Call WriteVar(UserFile, "INIT", "Y", val(UserList(userindex).POS.Y))

Call WriteVar(UserFile, "STATS", "GLD", val(UserList(userindex).Stats.GLD))
Call WriteVar(UserFile, "STATS", "BANCO", val(UserList(userindex).Stats.Banco))

Call WriteVar(UserFile, "STATS", "MaxHP", val(UserList(userindex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", val(UserList(userindex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "MaxMAN", val(UserList(userindex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MinMAN", val(UserList(userindex).Stats.MinMAN))
'
Call WriteVar(UserFile, "STATS", "MinSTA", val(UserList(userindex).Stats.MinSta))

Call WriteVar(UserFile, "STATS", "MaxHIT", val(UserList(userindex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", val(UserList(userindex).Stats.MinHIT))

Call WriteVar(UserFile, "STATS", "MaxAGU", val(UserList(userindex).Stats.MaxAGU))
Call WriteVar(UserFile, "STATS", "MinAGU", val(UserList(userindex).Stats.MinAGU))

Call WriteVar(UserFile, "STATS", "MaxHAM", val(UserList(userindex).Stats.MaxHam))
Call WriteVar(UserFile, "STATS", "MinHAM", val(UserList(userindex).Stats.MinHam))

Call WriteVar(UserFile, "STATS", "SkillPtsLibres", val(UserList(userindex).Stats.SkillPts))

Call WriteVar(UserFile, "STATS", "EXP", val(UserList(userindex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "ELV", val(UserList(userindex).Stats.ELV))

Call WriteVar(UserFile, "STATS", "NpcsMuertes", val(UserList(userindex).Stats.NPCsMuertos))

Call WriteVar(UserFile, "STATS", "MURIO", val(UserList(userindex).Stats.VecesMurioUsuario))


For LoopC = 1 To 3
Call WriteVar(UserFile, "STATS", "Recompensa" & LoopC, val(UserList(userindex).Recompensas(LoopC)))
Next LoopC

Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    'Call WriteVar(UserFile, "BANCO", "OBJ" & loopd, val(UserList(userindex).BancoInvent.Object(loopd).OBJIndex))
    'Call WriteVar(UserFile, "BANCO", "CANT" & loopd, val(UserList(userindex).BancoInvent.Object(loopd).Amount))
Call WriteVar(UserFile, "BANCO", "OBJ" & loopd, val(UserList(userindex).BancoInvent.Object(loopd).OBJIndex) & "-" & val(UserList(userindex).BancoInvent.Object(loopd).Amount))
Next loopd

Dim cad As String
For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(userindex).Stats.UserHechizos(LoopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
Next LoopC


For LoopC = 1 To MAX_INVENTORY_SLOTS
    'Call WriteVar(UserFile, "Invent", "OBJ" & LoopC, val(UserList(userindex).Invent.Object(LoopC).OBJIndex))
    'Call WriteVar(UserFile, "Invent", "CANT" & LoopC, val(UserList(userindex).Invent.Object(LoopC).Amount))
Call WriteVar(UserFile, "Invent", "OBJ" & LoopC, val(UserList(userindex).Invent.Object(LoopC).OBJIndex) & "-" & val(UserList(userindex).Invent.Object(LoopC).Amount))
Next LoopC



Call WriteVar(UserFile, "Invent", "CASCOSLOT", val(UserList(userindex).Invent.CascoEqpSlot))
Call WriteVar(UserFile, "Invent", "ARMORSLOT", val(UserList(userindex).Invent.ArmourEqpSlot))
Call WriteVar(UserFile, "Invent", "SHIELDSLOT", val(UserList(userindex).Invent.EscudoEqpSlot))
Call WriteVar(UserFile, "Invent", "WEAPONSLOT", val(UserList(userindex).Invent.WeaponEqpSlot))
Call WriteVar(UserFile, "Invent", "HERRAMIENTASLOT", val(UserList(userindex).Invent.HerramientaEqpslot))
Call WriteVar(UserFile, "Invent", "MUNICIONSLOT", val(UserList(userindex).Invent.MunicionEqpSlot))
Call WriteVar(UserFile, "Invent", "BARCOSLOT", val(UserList(userindex).Invent.BarcoSlot))


Call RevisarTops(userindex)

Exit Sub
errhandler:
Call LogError("Error en SaveUser")

End Sub

Sub LoadUserInit(userindex As Integer, UserFile As String)
On Error GoTo PErR


Dim LoopC As Integer
Dim Cantidad As Long

Dim Actual As String

With UserList(userindex)

If Len(UserFile) = 0 Then
Call CloseSocket(userindex)
Exit Sub
End If


'Actual = "facciones"
'Actual = UCase$(Actual)

Actual = "flags"

Dim Sistema As Byte

Sistema = val(GetVar(UserFile, UCase$(Actual), "SistemaNuevo"))
.flags.Ban = val(GetVar(UserFile, UCase$(Actual), "Ban"))
.flags.Navegando = val(GetVar(UserFile, UCase$(Actual), "Navegando"))
.flags.Envenenado = val(GetVar(UserFile, UCase$(Actual), "Envenenado"))
.flags.Silenciado = val(GetVar(UserFile, UCase$(Actual), "Silenciado"))
.Counters.TiempoPena = val(GetVar(UserFile, UCase$(Actual), "Pena")) * 60
.Counters.PenaMinar = val(GetVar(UserFile, UCase$(Actual), "PenaM"))
.flags.Denuncias = val(GetVar(UserFile, UCase$(Actual), "DenunciasCheat"))
.flags.DenunciasInsultos = val(GetVar(UserFile, UCase$(Actual), "DenunciasInsulto"))
.Password = GetVar(UserFile, UCase$(Actual), "Password")


Actual = "FACCIONES"
.Faccion.Bando = val(GetVar(UserFile, Actual, "Bando"))
.Faccion.BandoOriginal = val(GetVar(UserFile, Actual, "BandoOriginal"))
.Faccion.Matados(0) = val(GetVar(UserFile, Actual, "matados0"))
.Faccion.Matados(1) = val(GetVar(UserFile, Actual, "matados1"))
.Faccion.Matados(2) = val(GetVar(UserFile, Actual, "matados2"))
.Faccion.Jerarquia = val(GetVar(UserFile, Actual, "Jerarquia"))
.Faccion.Ataco(1) = val(GetVar(UserFile, Actual, "Ataco1"))
.Faccion.Ataco(2) = val(GetVar(UserFile, Actual, "Ataco2"))
.Faccion.Quests = val(GetVar(UserFile, Actual, "Quests"))
.Faccion.Torneos = val(GetVar(UserFile, Actual, "Torneos"))
.flags.ConsejoCiuda = val(GetVar(UserFile, Actual, "ConsejoCiuda"))
.flags.ConsejoCaoz = val(GetVar(UserFile, Actual, "ConsejoCaoz"))
.flags.AyudanteCiuda = val(GetVar(UserFile, Actual, "AyudanteCiuda"))
.flags.AyudanteCaoz = val(GetVar(UserFile, Actual, "AyudanteCaoz"))
.flags.MatadasenR = val(GetVar(UserFile, Actual, "RGanados"))
.flags.PerdidasenR = val(GetVar(UserFile, Actual, "RPerdidos"))


If Not ModoQuest And UserList(userindex).Faccion.Bando <> Neutral And UserList(userindex).Faccion.Bando <> UserList(userindex).Faccion.BandoOriginal Then UserList(userindex).Faccion.Bando = Neutral
    

Actual = "Guild"
UserList(userindex).GuildInfo.EsGuildLeader = val(GetVar(UserFile, UCase$(Actual), "EsGuildLeader"))
UserList(userindex).GuildInfo.echadas = val(GetVar(UserFile, UCase$(Actual), "Echadas"))
UserList(userindex).GuildInfo.Solicitudes = val(GetVar(UserFile, UCase$(Actual), "Solicitudes"))
UserList(userindex).GuildInfo.SolicitudesRechazadas = val(GetVar(UserFile, UCase$(Actual), "SolicitudesRechazadas"))
UserList(userindex).GuildInfo.VecesFueGuildLeader = val(GetVar(UserFile, UCase$(Actual), "VecesFueGuildLeader"))
UserList(userindex).GuildInfo.YaVoto = val(GetVar(UserFile, UCase$(Actual), "YaVoto"))
UserList(userindex).GuildInfo.FundoClan = val(GetVar(UserFile, UCase$(Actual), "FundoClan"))
UserList(userindex).GuildInfo.GuildName = GetVar(UserFile, UCase$(Actual), "GuildName")
UserList(userindex).GuildInfo.ClanFundado = GetVar(UserFile, UCase$(Actual), "ClanFundado")
UserList(userindex).GuildInfo.ClanesParticipo = val(GetVar(UserFile, UCase$(Actual), "ClanesParticipo"))
UserList(userindex).GuildInfo.GuildPoints = val(GetVar(UserFile, UCase$(Actual), "GuildPts"))



For LoopC = 1 To NUMATRIBUTOS
 .Stats.UserAtributos(LoopC) = GetVar(UserFile, "ATRIBUTOS", "AT" & LoopC)
 .Stats.UserAtributosBackUP(LoopC) = .Stats.UserAtributos(LoopC)
Next LoopC




For LoopC = 1 To NUMSKILLS
.Stats.UserSkills(LoopC) = val(GetVar(UserFile, "SKILLS", "SK" & LoopC))
Next LoopC



If Sistema Then
        Dim BancBoston As String
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        BancBoston = GetVar(UserFile, "BANCO", "OBJ" & LoopC)
            .BancoInvent.Object(LoopC).OBJIndex = ReadField$(1, BancBoston, 45)
            .BancoInvent.Object(LoopC).Amount = ReadField$(2, BancBoston, 45)
        Next LoopC
    
    
        For LoopC = 1 To MAX_INVENTORY_SLOTS
        BancBoston = GetVar(UserFile, "Invent", "OBJ" & LoopC)
            UserList(userindex).Invent.Object(LoopC).OBJIndex = ReadField$(1, BancBoston, 45)
            UserList(userindex).Invent.Object(LoopC).Amount = ReadField$(2, BancBoston, 45)
        Next LoopC
    
Else

        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            .BancoInvent.Object(LoopC).OBJIndex = GetVar(UserFile, "BANCO", "OBJ" & LoopC)
            .BancoInvent.Object(LoopC).Amount = GetVar(UserFile, "BANCO", "CANT" & LoopC)
        Next LoopC
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            UserList(userindex).Invent.Object(LoopC).OBJIndex = GetVar(UserFile, "Invent", "OBJ" & LoopC)
            UserList(userindex).Invent.Object(LoopC).Amount = GetVar(UserFile, "Invent", "CANT" & LoopC)
        Next LoopC


End If




.Invent.CascoEqpSlot = val(GetVar(UserFile, "Invent", "CASCOSLOT"))
.Invent.ArmourEqpSlot = val(GetVar(UserFile, "Invent", "ARMORSLOT"))
.Invent.EscudoEqpSlot = val(GetVar(UserFile, "Invent", "SHIELDSLOT"))
.Invent.WeaponEqpSlot = val(GetVar(UserFile, "Invent", "WEAPONSLOT"))
.Invent.HerramientaEqpslot = val(GetVar(UserFile, "Invent", "HERRAMIENTASLOT"))
.Invent.MunicionEqpSlot = val(GetVar(UserFile, "Invent", "MUNICIONSLOT"))
.Invent.BarcoSlot = val(GetVar(UserFile, "Invent", "BarcoSlot"))


For LoopC = 1 To MAXUSERHECHIZOS
.Stats.UserHechizos(LoopC) = val(GetVar(UserFile, "Hechizos", "H" & LoopC))
Next LoopC



.Stats.GLD = val(GetVar(UserFile, "STATS", "GLD"))
.Stats.Banco = val(GetVar(UserFile, "STATS", "BANCO"))
.Stats.MaxHP = val(GetVar(UserFile, "STATS", "MaxHP"))
.Stats.MinHP = val(GetVar(UserFile, "STATS", "MinHP"))
.Stats.MinSta = val(GetVar(UserFile, "STATS", "MinSTA"))
.Stats.MaxMAN = val(GetVar(UserFile, "STATS", "MaxMAN"))
.Stats.MinMAN = val(GetVar(UserFile, "STATS", "MinMAN"))
.Stats.MaxHIT = val(GetVar(UserFile, "STATS", "MaxHIT"))
.Stats.MinHIT = val(GetVar(UserFile, "STATS", "MinHIT"))
.Stats.MinAGU = val(GetVar(UserFile, "STATS", "MinAGU"))
.Stats.MinHam = val(GetVar(UserFile, "STATS", "MinHAM"))
.Stats.SkillPts = val(GetVar(UserFile, "STATS", "SkillPtsLibres"))
.Stats.Exp = val(GetVar(UserFile, "STATS", "EXP"))
.Stats.ELV = val(GetVar(UserFile, "STATS", "ELV"))
.Stats.ELU = ELUs(.Stats.ELV) 'val(GetVar(UserFile, "STATS", "ELU"))

.Stats.NPCsMuertos = val(GetVar(UserFile, "STATS", "NpcsMuertes"))
.Stats.VecesMurioUsuario = val(GetVar(UserFile, "STATS", "MURIO"))


For LoopC = 1 To 3
.Recompensas(LoopC) = GetVar(UserFile, "STATS", "Recompensa" & LoopC)
Next LoopC

    If .Stats.MinAGU < 1 Then .flags.Sed = 1
    If .Stats.MinHam < 1 Then .flags.Hambre = 1
    If .Stats.MinHP < 1 Then .flags.Muerto = 1



.Email = GetVar(UserFile, "INIT", "Email")
.Genero = GetVar(UserFile, "INIT", "Genero")
.Raza = GetVar(UserFile, "INIT", "Raza")
.Hogar = GetVar(UserFile, "INIT", "Hogar")
.Clase = GetVar(UserFile, "INIT", "Clase")

.POS.Map = val(GetVar(UserFile, "INIT", "Mapa"))
.POS.x = val(GetVar(UserFile, "INIT", "X"))
.POS.Y = val(GetVar(UserFile, "INIT", "Y"))




.OrigChar.Head = val(GetVar(UserFile, "INIT", "Head"))

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
    If Len(.Desc) >= 80 Then .Desc = Left$(.Desc, 80)
    
    
    If .Counters.TiempoPena > 0 Then
        .flags.Encarcelado = 1
        .Counters.Pena = Timer
    End If
    
    .Stats.MaxAGU = 100
    .Stats.MaxHam = 100
    Call CalcularSta(userindex)
    
    
    
.Desc = GetVar(UserFile, "INIT", "Descripcion")
.Moti = GetVar(UserFile, "INIT", "Motivo")

'OKEY XD

End With

Exit Sub
PErR:
Call LogError("Error en LoadUserInit User:" & UserList(userindex).Name & "  Desc." & Err.Description)
End Sub




'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal userindex As Integer)
If UserList(userindex).ConnID <> -1 Then 'And UserList(UserIndex).ConnIDvalida Then
    Call WSApiCloseSocket(UserList(userindex).ConnID)
    Call BorraSlotSock(UserList(userindex).ConnID)
    UserList(userindex).ConnIDvalida = False
End If
End Sub
Public Sub BorraSlotSock(ByVal Sock As Long, Optional ByVal CacheIndice As Long)
Dim Cant As Long
Cant = WSAPISock2Usr.Count
On Error Resume Next
WSAPISock2Usr.Remove CStr(Sock)
End Sub




Public Function EsChitUser(Procesos As String) As Boolean
'FIXIT: Declare 'X' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE
Dim x, i As Integer
Procesos = UCase$(Procesos)
    For i = 1 To Len(Procesos)
        For x = 1 To UBound(PalabrasP)
                If Mid$(Procesos, i, Len(PalabrasP(x))) = PalabrasP(x) And PalabrasP(x) <> "" Then
                EsChitUser = True
                Exit Function
                End If
            DoEvents
        Next x
    Next i


End Function


