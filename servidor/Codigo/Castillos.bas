Attribute VB_Name = "Castillos"
Public Const TIEMPOLIBRE = 5
Public DominadoPor As Integer
Public CrimiElejido As String
Public CiudaElejido As String
Public TiempoCast As Integer
Public ComenzoCuenta As Boolean
Public ComenzoTorneo As Boolean
Public Castillo As WorldPos
Public CastilloSala As WorldPos
Public ClaveCastillo As Integer
Public Sub RebisarVotos()
On Error GoTo erra:
Dim LoopC As Integer
Dim MayorCri As Integer
Dim QuienMayorCri As String
Dim MayorCiu As Integer
Dim QuienMayorCiu As String
For LoopC = 1 To LastUser
    If (UserList(LoopC).NumVotos > IIf(Criminal(LoopC), MayorCri, MayorCiu)) Then
        If Criminal(LoopC) Then
            MayorCri = UserList(LoopC).NumVotos
            QuienMayorCri = UserList(LoopC).Name
        Else
            MayorCiu = UserList(LoopC).NumVotos
            QuienMayorCiu = UserList(LoopC).Name
        End If
    End If
    UserList(LoopC).Voto = False
    UserList(LoopC).NumVotos = 0
Next LoopC
CrimiElejido = QuienMayorCri
CiudaElejido = QuienMayorCiu
Exit Sub
erra:
Call LogError("Error en RevisarVotos")
End Sub
Public Sub ComenzarDuelo()
On Error GoTo erra:
Dim LoopC As Integer
Dim Name As String
ComenzoTorneo = True
For LoopC = 1 To 2
    Name = IIf(LoopC = 1, CiudaElejido, CrimiElejido)
    If NameIndex(Name) > 0 Then
        If UserList(NameIndex(Name)).flags.Montado = 1 Then Call Desmontar(NameIndex(Name))
        If UserList(NameIndex(Name)).flags.Muerto = 1 Then Call RevivirUsuario(NameIndex(Name))
        Call WarpUserChar(NameIndex(Name), CastilloSala.Map, CastilloSala.x - 1 + LoopC, CastilloSala.y, True)
    Else
        Call SendData(ToAll, 0, 0, "||" & Name & " (Representante de " & IIf(LoopC = 1, "Ciudadanos", "Criminales") & ") no se presento a la batalla por el castillo." & FONTTYPE_GUILD)
        Call SendData(ToAll, 0, 0, "||El Castillo queda en manos de los " & IIf(LoopC = 2, "Ciudadanos", "Criminales") & FONTTYPE_GUILD)
        DominadoPor = IIf(LoopC = 1, 2, 1)
        ResetCastilloInfo
    End If
Next LoopC
Exit Sub
erra:
Call LogError("Error en ComenzarDuelo")
End Sub
Public Sub ResetCastilloInfo()
CrimiElejido = ""
CiudaElejido = ""
ComenzoCuenta = False
ComenzoTorneo = False
TiempoCast = 0
End Sub
Public Function CastilloDuelo(Name As String) As Boolean
CastilloDuelo = False
If UCase$(Name) = UCase$(CrimiElejido) Or UCase$(Name) = UCase$(CiudaElejido) Then CastilloDuelo = True
End Function
Public Function PuedeEntrarCastillo(UserIndex As Integer) As Boolean
On Error GoTo erra:
Dim Name As String
Name = UserList(UserIndex).Name
PuedeEntrarCastillo = False
If UCase$(Name) = UCase$(CiudaElejido) Or UCase$(Name) = UCase$(CrimiElejido) Then
    PuedeEntrarCastillo = True
    Exit Function
End If
If ComenzoTorneo Then Exit Function
If DominadoPor = 2 And Criminal(UserIndex) Then
    PuedeEntrarCastillo = True
    Exit Function
ElseIf (DominadoPor = 1) And (Not Criminal(UserIndex)) Then
    PuedeEntrarCastillo = True
    Exit Function
End If
Exit Function
erra:
Call LogError("Error en PuedeEntrarCatillo")
End Function
Public Sub DesalojarCastillo()
On Error GoTo erra:
Dim LoopC As Integer
For LoopC = 1 To LastUser
    If UserList(LoopC).Pos.Map = Castillo.Map Then
        Call WarpUserChar(LoopC, 1, 50, 50, True)
    End If
Next LoopC
Exit Sub
erra:
Call LogError("Error en DesalojarCastillo")
End Sub
