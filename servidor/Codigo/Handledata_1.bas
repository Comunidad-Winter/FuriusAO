Attribute VB_Name = "Handledata_1"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Public Sub HandleData1(userindex As Integer, rdata As String, Procesado As Boolean)
Dim tInt As Integer, tIndex As Integer, x As Integer, Y As Integer
Dim Arg1 As String, Arg2 As String, arg3 As String
Dim nPos As WorldPos
Dim tLong As Long
'FIXIT: Declare 'ind' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim ind

Procesado = True

Select Case UCase$(Left$(rdata, 1))
    Case "\"
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        
        rdata = Right$(rdata, Len(rdata) - 1)
        tName = ReadField(1, rdata, 32)
        tIndex = NameIndex(tName)
        
        If tIndex <> 0 Then
            If UserList(tIndex).flags.Muerto = 1 Then Exit Sub
    
            If Len(rdata) <> Len(tName) Then
                tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
            Else
                tMessage = " "
            End If
             
            If Not EnPantalla(UserList(userindex).POS, UserList(tIndex).POS, 1) Then
                Call SendData(ToIndex, userindex, 0, "2E")
                Exit Sub
            End If
             
            ind = UserList(userindex).Char.CharIndex
             
            If InStr(tMessage, "°") Then Exit Sub
    
            If UserList(tIndex).flags.Privilegios > 0 And UserList(userindex).flags.Privilegios = 0 Then
                Call SendData(ToIndex, userindex, 0, "3E")
                Exit Sub
            End If
    
            Call SendData(ToIndex, userindex, UserList(userindex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Call SendData(ToIndex, tIndex, UserList(userindex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Call SendData(ToGMArea, userindex, UserList(userindex).POS.Map, "||" & vbCyan & "°" & tMessage & "°" & str(ind))
            Exit Sub
        End If
        
        Call SendData(ToIndex, userindex, 0, "3E")
        Exit Sub
            
    Case ";"

        Dim Modo As String
        
        rdata = Right$(rdata, Len(rdata) - 1)
        If Right$(rdata, Len(rdata) - 1) = " " Or Right$(rdata, Len(rdata) - 1) = "-" Then rdata = "1 "
        If Len(rdata) = 1 Then Exit Sub
       
        
        Modo = Left$(rdata, 1)
        
    If UserList(userindex).flags.Silenciado > 0 Then
    Call SendData(ToIndex, userindex, 0, "||Te encuentras silenciado por " & UserList(userindex).flags.Silenciado & " minutos." & FONTTYPE_VERDE)
   rdata = "1Estoy silenciado por " & UserList(userindex).flags.Silenciado & " minutos."
   ' Exit Sub
    End If
    
    
    
        rdata = Replace(Right$(rdata, Len(rdata) - 1), "~", "-")
        
    Select Case Modo
            
        Case 1, 4, 5
            
            If InStr(rdata, "°") Then Exit Sub
            
            If (Modo = 4 Or Modo = 5) And UserList(userindex).flags.Muerto = 1 And UserList(userindex).Clase <> CLERIGO Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
            End If
            
            If UserList(userindex).flags.Privilegios = 1 Then Call LogGM(UserList(userindex).Name, "Dijo: " & rdata, True)
            If InStr(1, rdata, Chr$(255)) Then rdata = Replace(rdata, Chr$(255), " ")
            
            ind = UserList(userindex).Char.CharIndex
            Dim Color As Long
            Dim IndexSendData As Byte
            
            If Modo = 4 Then
                Color = vbRed
            ElseIf Modo = 5 Then
                Color = vbGreen
            ElseIf UserList(userindex).flags.Privilegios Then
                Color = &H80FF&
            ElseIf UserList(userindex).flags.Quest And UserList(userindex).Faccion.Bando <> Neutral Then
                If UserList(userindex).Faccion.Bando = Real Then
                    Color = vbBlue
                Else: Color = vbRed
                End If
            ElseIf UserList(userindex).flags.Muerto Then
                Color = vbYellow
            Else: Color = vbWhite
            End If
    
            If UserList(userindex).flags.Privilegios > 0 Or UserList(userindex).Clase = CLERIGO Then
                IndexSendData = ToPCArea
            ElseIf UserList(userindex).flags.Muerto Then
                IndexSendData = ToMuertos
            Else
                IndexSendData = ToPCAreaVivos
            End If
   
            If Modo = 5 Then rdata = "* " & rdata & " *"

            Call SendData(IndexSendData, userindex, UserList(userindex).POS.Map, "||" & Color & "°" & rdata & "°" & str(ind))
            Exit Sub
            
        Case 2
            
            If UserList(userindex).flags.Muerto Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
            End If
            
            tIndex = UserList(userindex).flags.Whispereando
            
            If tIndex Then
                If UserList(tIndex).flags.Muerto Then Exit Sub
    
                If Not EnPantalla(UserList(userindex).POS, UserList(tIndex).POS, 1) Then
                    Call SendData(ToIndex, userindex, 0, "2E")
                    Exit Sub
                End If
                
                ind = UserList(userindex).Char.CharIndex
                
                If InStr(rdata, "°") Then Exit Sub

                If UserList(tIndex).flags.Privilegios > 0 And UserList(tIndex).flags.AdminInvisible Then
                    Call SendData(ToIndex, userindex, 0, "3E")
                    Call SendData(ToIndex, tIndex, UserList(userindex).POS.Map, "||" & vbBlue & "°" & rdata & "°" & str(ind))
                    Exit Sub
                End If
                
                If UserList(userindex).flags.Privilegios = 1 Then Call LogGM(UserList(userindex).Name, "Grito: " & rdata, True)
                
                If EnPantalla(UserList(userindex).POS, UserList(tIndex).POS, 1) Then
                    Call SendData(ToIndex, userindex, 0, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                    Call SendData(ToIndex, tIndex, 0, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                    Call SendData(ToGMArea, userindex, UserList(userindex).POS.Map, "||" & vbCyan & "°" & rdata & "°" & str(ind))
                Else
                    Call SendData(ToIndex, userindex, 0, "{F")
                    UserList(userindex).flags.Whispereando = 0
                End If
            End If
            
            Exit Sub
        
        Case 3
            If UserList(userindex).flags.Muerto And UserList(userindex).Clase <> CLERIGO Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
           End If
        
            If Len(rdata) And Len(UserList(userindex).GuildInfo.GuildName) > 0 Then Call SendData(ToDiosesYclan, userindex, 0, "::" & UserList(userindex).Name & "> " & rdata)

        Exit Sub
        
        Case 6
            If UserList(userindex).flags.Party = 0 Then Exit Sub
              If UserList(userindex).flags.Muerto Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
        End If
        
            If Len(rdata) > 0 Then
                Call SendData(ToParty, userindex, 0, "||" & UserList(userindex).Name & ": " & rdata & FONTTYPE_PARTY)
            End If
            Exit Sub
                
        Case 7
            If UserList(userindex).flags.Privilegios = 0 Then Exit Sub
            
            Call LogGM(UserList(userindex).Name, "Mensaje a Gms:" & rdata, (UserList(userindex).flags.Privilegios = 1))
            If Len(rdata) > 0 Then
                Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & "> " & rdata & "~255~255~255~0~1")
            End If
            
            Exit Sub
    
        End Select
        
    Case "M"
        Dim Mide As Double
        rdata = Right$(rdata, Len(rdata) - 1)

        If UserList(userindex).flags.Trabajando Then

                Call SacarModoTrabajo(userindex)

        End If
        If UserList(userindex).flags.Bovediando = 1 Then Call SendData(ToIndex, userindex, 0, "PU" & UserList(userindex).POS.x & "," & UserList(userindex).POS.Y): Exit Sub
        If Not UserList(userindex).flags.Descansar And Not UserList(userindex).flags.Meditando _
           And UserList(userindex).flags.Paralizado = 0 Then
            Call MoveUserChar(userindex, val(rdata))
        ElseIf UserList(userindex).flags.Descansar Then
            UserList(userindex).flags.Descansar = False
            Call SendData(ToIndex, userindex, 0, "DOK")
            Call SendData(ToIndex, userindex, 0, "DN")
            Call MoveUserChar(userindex, val(rdata))
        End If

        If UserList(userindex).flags.Oculto Then
            If Not (UserList(userindex).Clase = LADRON And UserList(userindex).Recompensas(2) = 1) Then
                UserList(userindex).flags.Oculto = 0
                UserList(userindex).flags.Invisible = 0
                Call SendData(ToMap, 0, UserList(userindex).POS.Map, ("V3" & UserList(userindex).Char.CharIndex & ",0"))
                Call SendData(ToIndex, userindex, 0, "V5")
            End If
        End If

        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 2))
Case "%1"
Call SendData(ToIndex, userindex, 0, "PONG")
Exit Sub
Case "FP"
Call SendData(ToAdmins, 0, 0, "||FPS de " & UserList(userindex).Name & ": " & Right$(rdata, Len(rdata) - 2) & FONTTYPE_CELESTE)
Exit Sub

Case "^~"
'If UserList(userindex).POS.Map <> MapaJuego Then Exit Sub ' aseguro de que este en el mapa de el juego'

'     If UserList(userindex).flags.TargetUser <= 0 Then
'          Call SendData(ToIndex, userindex, 0, "||No tocaste a nadie " & FONTTYPE_INFO)
'            Exit Sub
'      End If'

'     If Distancia(UserList(userindex).POS, UserList(UserList(userindex).flags.TargetUser).POS) > 2 Then
'          Call SendData(ToIndex, userindex, 0, "||Haaay!, por poco, pero no llegaste y te esquivo" & FONTTYPE_INFO)
 '         Exit Sub
 '     End If

'i f UserList(userindex).Name = UserList(UserList(userindex).flags.TargetUser).Name Then
'        Call SendData(ToIndex, userindex, 0, "||O.o Te tocaste a vos mismo" & FONTTYPE_INFO)
'        Exit Sub
'End If

'If UserList(UserList(userindex).flags.TargetUser).POS.Map <> MapaJuego Then  ' ultima verificación, si clickeo a alguien antes de entrar al mapa, que no lo convierta en mancha

 '        Call SendData(ToIndex, userindex, 0, "||Toca a alguien en tu mismo mapa" & FONTTYPE_INFO)
 '        Exit Sub
'End If

'Call HacerMancha(userindex, UserList(userindex).flags.TargetUser)
 '  Exit Sub

    Case "ZI"
        rdata = Right$(rdata, Len(rdata) - 2)
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
        Dim Bait(1 To 2) As Byte
        Bait(1) = val(ReadField(1, rdata, 44))
        Bait(2) = val(ReadField(2, rdata, 44))
        
        Select Case Bait(2)
            Case 0
                Bait(2) = Bait(1) - 1
            Case 1
                Bait(2) = Bait(1) + 1
            Case 2
                Bait(2) = Bait(1) - 5
            Case 3
                Bait(2) = Bait(1) + 5
        End Select
        
        If Bait(2) > 0 And Bait(2) <= MAX_INVENTORY_SLOTS Then Call AcomodarItems(userindex, Bait(1), Bait(2))
        
        Exit Sub
    Case "TI"
        If UserList(userindex).flags.Navegando = 1 Or _
           UserList(userindex).flags.Muerto = 1 Or _
                          UserList(userindex).flags.EnDM Then Exit Sub
           
        If UserList(userindex).flags.Privilegios = 1 Then Exit Sub
        If UserList(userindex).POS.Map = 66 Then Exit Sub
        If UserList(userindex).POS.Map = 194 And UserList(userindex).flags.Privilegios = 0 Then Exit Sub
        If UserList(userindex).POS.Map = 196 And UserList(userindex).flags.Privilegios = 0 Then Exit Sub
        
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If val(Arg1) = FLAGORO Then
            Call TirarOro(val(Arg2), userindex)
            Call SendUserORO(userindex)
            Exit Sub
        Else
            If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) Then
                If UserList(userindex).Invent.Object(val(Arg1)).OBJIndex = 0 Then
                        Exit Sub
                End If
                Call DropObj(userindex, val(Arg1), val(Arg2), UserList(userindex).POS.Map, UserList(userindex).POS.x, UserList(userindex).POS.Y)
            Else
                Exit Sub
            End If
        End If
        Exit Sub
    Case "SF"
        rdata = Right$(rdata, Len(rdata) - 2)
        If Not PuedeFaccion(userindex) Then Exit Sub
        If UserList(userindex).Faccion.BandoOriginal Then Exit Sub
        tInt = val(rdata)
        
        If tInt = Neutral Then
            If UserList(userindex).Faccion.Bando <> Neutral Then
                Call SendData(ToIndex, userindex, 0, "7&")
            Else: Call SendData(ToIndex, userindex, 0, "0&")
            End If
            Exit Sub
        End If
        
        If UserList(userindex).Faccion.Matados(tInt) > UserList(userindex).Faccion.Matados(Enemigo(tInt)) Then
            Call SendData(ToIndex, userindex, 0, Mensajes(tInt, 9))
            Exit Sub
        End If
        
        Call SendData(ToIndex, userindex, 0, Mensajes(tInt, 10))
        UserList(userindex).Faccion.BandoOriginal = tInt
        UserList(userindex).Faccion.Bando = tInt
        UserList(userindex).Faccion.Ataco(tInt) = 0
        If Not PuedeFaccion(userindex) Then Call SendData(ToIndex, userindex, 0, "SUFA0")
        
        Call UpdateUserChar(userindex)
        
        Exit Sub
    Case "LH"
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 2)
        UserList(userindex).flags.Hechizo = val(rdata)
        Exit Sub
    Case "WH"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        x = CInt(Arg1)
        Y = CInt(Arg2)
        If Not InMapBounds(x, Y) Then Exit Sub
        Call LookatTile(userindex, UserList(userindex).POS.Map, x, Y)
        
        If UserList(userindex).flags.TargetUser = userindex Then
            Call SendData(ToIndex, userindex, 0, "{C")
            Exit Sub
        End If
        
        If UserList(userindex).flags.TargetUser Then
            UserList(userindex).flags.Whispereando = UserList(userindex).flags.TargetUser
            Call SendData(ToIndex, userindex, 0, "{B" & UserList(UserList(userindex).flags.Whispereando).Name)
        Else
            Call SendData(ToIndex, userindex, 0, "{D")
        End If
        
        Exit Sub
    Case "LC"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        Dim POS As WorldPos
        POS.Map = UserList(userindex).POS.Map
        POS.x = CInt(Arg1)
        POS.Y = CInt(Arg2)
        If Not EnPantalla(UserList(userindex).POS, POS, 1) Then Exit Sub
        Call LookatTile(userindex, UserList(userindex).POS.Map, POS.x, POS.Y)
        Exit Sub
    Case "RC"
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        x = CInt(Arg1)
        Y = CInt(Arg2)
        Call Accion(userindex, UserList(userindex).POS.Map, x, Y)
        Exit Sub
  
    Case "UK"
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If

        rdata = Right$(rdata, Len(rdata) - 2)
        Select Case val(rdata)
            Case Robar
                Call SendData(ToIndex, userindex, 0, "T01" & Robar)
            Case Magia
                Call SendData(ToIndex, userindex, 0, "T01" & Magia)
            Case Domar
                Call SendData(ToIndex, userindex, 0, "T01" & Domar)
            Case Invitar
                Call SendData(ToIndex, userindex, 0, "T01" & Invitar)
                
            Case Ocultarse
                
                If UserList(userindex).flags.Navegando Then
                      Call SendData(ToIndex, userindex, 0, "6E")
                      Exit Sub
                End If
                
                If UserList(userindex).flags.Oculto Then
                      Call SendData(ToIndex, userindex, 0, "7E")
                      Exit Sub
                End If
                
                Call DoOcultarse(userindex)
        End Select
        Exit Sub
        'LEITO
     Case "CS" 'Recibe consulta para GM
        Consulta = Right$(rdata, Len(rdata) - 2)
        'Aquí declara que la variable del nombre de personaje que se envía desde el cliente
        'contiene la consulta en sí, así es más fácil de consultarla, menos liosa
        'UserList(UserIndex).Name = Right$(rdata, Len(rdata) - 1 - Len(ReadField(1, rdata, 44)))
        'Donde ReadField(1, rdata, 44) es el nombre de quien ha enviado el SOS, y, por la otra parte
        'está la consulta
        If Not Ayuda.Existe(UserList(userindex).Name) And Not Consultas.Existe(Consulta) Then
       '     Call SendData(ToIndex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
            Call Ayuda.Push(rdata, UserList(userindex).Name)
            Call Consultas.Push(rdata, Consulta)
            Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " ha enviado un soporte" & FONTTYPE_furius)
        Else
            Call Ayuda.Quitar(UserList(userindex).Name)
            Call Ayuda.Push(rdata, UserList(userindex).Name)
            Call Consultas.Quitar(Consulta)
            Call Consultas.Push(rdata, Consulta)
          '  Call SendData(ToIndex, userindex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
        End If
        Exit Sub
    Case "CJ" 'El GM texto de la respuesta
        Restext = Right$(rdata, Len(rdata) - 2)
        Call AccionRespuestaGm(Resnick, Restext)
        Exit Sub
    Case "RR" 'El GM nick de la respuesta
        Resnick = Right$(rdata, Len(rdata) - 2)
        Exit Sub
        'LEITOf
End Select

Select Case UCase$(rdata)
    Case "SOS"
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
       Else
    Call SendData(ToGuildMembers, userindex, 0, "||" & UserList(userindex).Name & "(" & UserList(userindex).Stats.MinHP & "/" & UserList(userindex).Stats.MaxHP & ") pide ayuda en " & UserList(userindex).POS.Map & " , " & UserList(userindex).POS.x & " , " & UserList(userindex).POS.Y & FONTTYPE_furius)
End If
    Case "RPU"
        Call SendData(ToIndex, userindex, 0, "PU" & UserList(userindex).POS.x & "," & UserList(userindex).POS.Y)
        Exit Sub
    Case "AT"
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
        End If
        If UserList(userindex).Invent.WeaponEqpObjIndex Then
            If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil Or ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Baculo Then
                Call SendData(ToIndex, userindex, 0, "||No puedes usar así esta arma." & FONTTYPE_FIGHT)
                Exit Sub
            End If
        End If
        
        Call UsuarioAtaca(userindex)
        
        Exit Sub
                '    Case "INFOMASCOTA"
                'Call SendMascBox(userindex)
            'Exit Sub
    Case "AG"
        If UserList(userindex).flags.Muerto Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
        End If
        
   
   
   
   
        Call GetObj(userindex)
        Exit Sub
   ' Case "SEGCLAN"
       ' If UserList(userindex).flags.Seguroclan Then
           '   Call SendData(ToIndex, userindex, 0, "1O")
     '  Else
         '     Call SendData(ToIndex, userindex, 0, "9K")
      '  End If
     '   UserList(userindex).flags.Seguroclan = Not UserList(userindex).flags.Seguroclan
       'Exit Sub
    Case "ATRI"
        Call EnviarAtrib(userindex)
        Exit Sub
    Case "FAMA"
        Call EnviarFama(userindex)
        Call EnviarMiniSt(userindex)
        Exit Sub
    Case "ESKI"
        Call EnviarSkills(userindex)
        Exit Sub
    Case "PARSAL"
        Dim i As Integer
        If UserList(userindex).flags.Party Then
            If Party(UserList(userindex).PartyIndex).NroMiembros = 2 Then
                Call RomperParty(userindex)
            Else: Call SacarDelParty(userindex)
            End If
        Else
            Call SendData(ToIndex, userindex, 0, "||No estás en party." & FONTTYPE_PARTY)
        End If
        Exit Sub
    Case "PARINF"
        Call EnviarIntegrantesParty(userindex)
        Exit Sub
    
    Case "FINCOM"
        
        UserList(userindex).flags.Comerciando = False
        Call SendData(ToIndex, userindex, 0, "FINCOMOK")
        Exit Sub
    Case "FINCOMUSU"
        If UserList(userindex).ComUsu.DestUsu > 0 Then
            If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                Call SendData(ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "6R" & UserList(userindex).Name)
                Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
            End If
        End If
        
        Call FinComerciarUsu(userindex)
        Exit Sub

    Case "FINBAN"
        UserList(userindex).flags.Comerciando = False
        Call SendData(ToIndex, userindex, 0, "FINBANOK")
        Exit Sub
        
    Case "FINTIE"
        UserList(userindex).flags.Comerciando = False
        Call SendData(ToIndex, userindex, 0, "FINTIEOK")
        Exit Sub

    Case "COMUSUOK"
        
        Call AceptarComercioUsu(userindex)
        Exit Sub
    Case "COMUSUNO"
        
        If UserList(userindex).ComUsu.DestUsu Then
            Call SendData(ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "7R" & UserList(userindex).Name)
            Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
        End If
        Call SendData(ToIndex, userindex, 0, "8R")
        Call FinComerciarUsu(userindex)
        Exit Sub
    Case "GLINFO"
        If UserList(userindex).GuildInfo.EsGuildLeader Then
            If UserList(userindex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, userindex, 0, "GINFIG")
            Else
                Call SendGuildLeaderInfo(userindex)
            End If
        ElseIf Len(UserList(userindex).GuildInfo.GuildName) > 0 Then
            If UserList(userindex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, userindex, 0, "GINFII")
            Else
                Call SendGuildsStats(userindex)
            End If
        Else
            If UserList(userindex).flags.InfoClanEstatica Then
                Call SendData(ToIndex, userindex, 0, "GINFIJ")
            Else: Call SendGuildsList(userindex)
            End If
        End If
        
        Exit Sub

End Select

 Select Case UCase$(Left$(rdata, 2))
 
    Case "NL"
    rdata = Right$(rdata, Len(rdata) - 2)
        If Len(rdata) > 0 Then
        rdata = DesencriptarFPS(rdata)
                If val(rdata) < 5 Then FPSBajos = True
        UserList(userindex).flags.Fps = rdata
        End If
 
    Case "(A"
        If PuedeDestrabarse(userindex) Then
            Call ClosestLegalPos(UserList(userindex).POS, nPos)
            If InMapBounds(nPos.x, nPos.Y) Then Call WarpUserChar(userindex, nPos.Map, nPos.x, nPos.Y, True)
        End If
        
        Exit Sub
    Case "GM"
        rdata = Right$(rdata, Len(rdata) - 2)
        Dim GMDia As String
        Dim GMMapa As String
        Dim GMPJ As String
        Dim GMMail As String
        Dim GMGM As String
        Dim GMTitulo As String
        Dim GMMensaje As String
        
        GMDia = Format(Now, "yyyy-mm-dd hh:mm:ss")
        GMMapa = UserList(userindex).POS.Map & " - " & UserList(userindex).POS.x & " - " & UserList(userindex).POS.Y
        GMPJ = UserList(userindex).Name
        GMMail = UserList(userindex).Email
        GMGM = ReadField(1, rdata, 172)
        GMTitulo = ReadField(2, rdata, 172)
        GMMensaje = ReadField(3, rdata, 172)
        conn.Execute "INSERT INTO sos(fecha,mapa,personaje,email,servidor,gm,asunto,mensaje,respondido,censura,old,respondidopor,respondidoel,respuesta) values(""" & GMDia & """,""" & GMMapa & """,""" & GMPJ & """,""" & GMMail & """, 1,""" & GMGM & """, """ & GMTitulo & """, """ & GMMensaje & """,0,0,0,0,0,0)"
        Call SendData(ToAdmins, 0, 9, "3B" & GMTitulo & "," & GMPJ)
  
        Exit Sub
        
    End Select
        
 Select Case UCase$(Left$(rdata, 3))
    
    Case "SH+"
       Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " está chiteando xD, es new" & FONTTYPE_INFO)
    
    Case "FRF"
        rdata = Right$(rdata, Len(rdata) - 3)
        For i = 1 To 10
            If UserList(userindex).flags.Espiado(i) > 0 Then
                If UserList(UserList(userindex).flags.Espiado(i)).flags.Privilegios > 1 Then Call SendData(ToIndex, UserList(userindex).flags.Espiado(i), 0, "{{" & UserList(userindex).Name & "," & rdata)
            End If
        Next
        Exit Sub
    Case "USA"
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
            If UserList(userindex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(userindex, val(rdata), 0)
        Exit Sub
    Case "USE"
        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
            If UserList(userindex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(userindex, val(rdata), 1)
        Exit Sub
    Case "CNS"
        Dim Arg5 As Integer
        rdata = Right$(rdata, Len(rdata) - 3)
        
        x = CInt(ReadField(1, rdata, 32))
        Arg5 = CInt(ReadField(2, rdata, 32))
        If Arg5 < 1 Then Exit Sub
        If x < 1 Then Exit Sub
        If ObjData(x).SkHerreria = 0 Then Exit Sub
        Call HerreroConstruirItem(userindex, x, val(Arg5))
        Exit Sub
        
    Case "CNC"
        rdata = Right$(rdata, Len(rdata) - 3)
        
        x = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If Arg1 < 1 Then Exit Sub
        If x < 1 Or ObjData(x).SkCarpinteria = 0 Then Exit Sub
        Call CarpinteroConstruirItem(userindex, x, val(Arg1))
        Exit Sub
        Case "CNA" ' Construye alquimia
        rdata = Right$(rdata, Len(rdata) - 3)
        x = ReadField(1, rdata, 44)
        If x < 1 Or ObjData(x).SkAlquimia = 0 Then Exit Sub
        Call AlquimiaConstruirItem(userindex, x, ReadField(2, rdata, 44))
        Exit Sub
    Case "SCR"
        rdata = Right$(rdata, Len(rdata) - 3)
        
        x = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If x < 1 Or ObjData(x).SkSastreria = 0 Then Exit Sub
        Call SastreConstruirItem(userindex, x, val(Arg1))
        Exit Sub
    
    Case "WLC"
        rdata = Right$(rdata, Len(rdata) - 3)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        arg3 = ReadField(3, rdata, 44)
        If Len(arg3) = 0 Or Len(Arg2) = 0 Or Len(Arg1) = 0 Then Exit Sub
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(arg3) Then Exit Sub
        
        POS.Map = UserList(userindex).POS.Map
        POS.x = CInt(Arg1)
        POS.Y = CInt(Arg2)
        tLong = CInt(arg3)
        
        If UserList(userindex).flags.Muerto = 1 Or _
           UserList(userindex).flags.Descansar Or _
           UserList(userindex).flags.Meditando Or _
           Not InMapBounds(POS.x, POS.Y) Then Exit Sub
        
        If Not EnPantalla(UserList(userindex).POS, POS, 1) Then
            Call SendData(ToIndex, userindex, 0, "PU" & UserList(userindex).POS.x & "," & UserList(userindex).POS.Y)
            Exit Sub
        End If
        
        Select Case tLong
        
        Case Proyectiles
            Dim TU As Integer, tN As Integer
            
            If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Or _
            UserList(userindex).Invent.MunicionEqpObjIndex = 0 Then Exit Sub
            
            If UserList(userindex).Invent.WeaponEqpSlot < 1 Or UserList(userindex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Or _
            UserList(userindex).Invent.MunicionEqpSlot < 1 Or UserList(userindex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Or _
            ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).ObjType <> OBJTYPE_FLECHAS Or _
            UserList(userindex).Invent.Object(UserList(userindex).Invent.MunicionEqpSlot).Amount < 1 Or _
            ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then Exit Sub
            
            If TiempoTranscurrido(UserList(userindex).Counters.LastFlecha) < IntervaloUserFlechas Then Exit Sub
            If TiempoTranscurrido(UserList(userindex).Counters.LastHechizo) < IntervaloUserPuedeHechiGolpe Then Exit Sub
            If TiempoTranscurrido(UserList(userindex).Counters.LastGolpe) < IntervaloUserPuedeAtacar Then Exit Sub
            
            UserList(userindex).Counters.LastFlecha = Timer
            Call SendData(ToIndex, userindex, 0, "LF")
            
            If UserList(userindex).Stats.MinSta >= 10 Then
                 Call QuitarSta(userindex, RandomNumber(1, 10))
            Else
                 Call SendData(ToIndex, userindex, 0, "9E")
                 Exit Sub
            End If
             
            Call LookatTile(userindex, UserList(userindex).POS.Map, val(Arg1), val(Arg2))
            
            TU = UserList(userindex).flags.TargetUser
            tN = UserList(userindex).flags.TargetNpc
                            
            If TU = userindex Then
                Call SendData(ToIndex, userindex, 0, "3N")
                Exit Sub
            End If

            Call QuitarUnItem(userindex, UserList(userindex).Invent.MunicionEqpSlot)
            Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW200")
            If UserList(userindex).Invent.MunicionEqpSlot Then
                UserList(userindex).Invent.Object(UserList(userindex).Invent.MunicionEqpSlot).Equipped = 1
                Call UpdateUserInv(False, userindex, UserList(userindex).Invent.MunicionEqpSlot)
            End If
            
            If tN Then
                If Npclist(tN).Attackable Then Call UsuarioAtacaNpc(userindex, tN)
            ElseIf TU Then
                If TU <> userindex Then
                    Call UsuarioAtacaUsuario(userindex, TU)
                    SendUserHP TU
                End If
            Else
                Call SendData(ToIndex, userindex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If
            
            
                
                
                
                
        Case Invitar
            Call LookatTile(userindex, UserList(userindex).POS.Map, POS.x, POS.Y)
            
            If UserList(userindex).flags.TargetUser = 0 Then
                Call SendData(ToIndex, userindex, 0, "||No hay nadie a quien invitar." & FONTTYPE_PARTY)
                Exit Sub
            End If
            
            If UserList(userindex).flags.Privilegios > 0 Or UserList(UserList(userindex).flags.TargetUser).flags.Privilegios > 0 Then Exit Sub

            Call DoInvitar(userindex, UserList(userindex).flags.TargetUser)
            
        Case Magia

            
            If UserList(userindex).flags.Privilegios = 1 Then Exit Sub
            
            Call LookatTile(userindex, UserList(userindex).POS.Map, POS.x, POS.Y)
            
            If UserList(userindex).flags.Hechizo Then
                Call LanzarHechizo(UserList(userindex).flags.Hechizo, userindex)
                UserList(userindex).flags.Hechizo = 0
            Else
                Call SendData(ToIndex, userindex, 0, "4N")
            End If
            
        Case Robar
               If TiempoTranscurrido(UserList(userindex).Counters.LastTrabajo) < 1 Then Exit Sub
               If MapInfo(UserList(userindex).POS.Map).Pk Or (UserList(userindex).Clase = LADRON) Then
               
                    
                    Call LookatTile(userindex, UserList(userindex).POS.Map, POS.x, POS.Y)

                    If UserList(userindex).flags.TargetUser > 0 And UserList(userindex).flags.TargetUser <> userindex Then
                       If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 0 Then
                            nPos.Map = UserList(userindex).POS.Map
                            nPos.x = POS.x
                            nPos.Y = POS.Y
                            
                            If Distancia(nPos, UserList(userindex).POS) > 4 Or (Not (UserList(userindex).Clase = LADRON And UserList(userindex).Recompensas(3) = 1) And Distancia(nPos, UserList(userindex).POS) > 2) Then
                                Call SendData(ToIndex, userindex, 0, "DL")
                                Exit Sub
                            End If

                            Call DoRobar(userindex, UserList(userindex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(ToIndex, userindex, 0, "4S")
                    End If
                Else
                    Call SendData(ToIndex, userindex, 0, "5S")
                End If
        Case Domar
          
          
          
          Dim CI As Integer
          
          Call LookatTile(userindex, UserList(userindex).POS.Map, POS.x, POS.Y)
          CI = UserList(userindex).flags.TargetNpc
          
          If CI Then
                   If Npclist(CI).flags.Domable Then
                        nPos.Map = UserList(userindex).POS.Map
                        nPos.x = POS.x
                        nPos.Y = POS.Y
                        If Distancia(nPos, Npclist(UserList(userindex).flags.TargetNpc).POS) > 2 Then
                              Call SendData(ToIndex, userindex, 0, "DL")
                              Exit Sub
                        End If
                        If Npclist(CI).flags.AttackedBy Then
                              Call SendData(ToIndex, userindex, 0, "7S")
                              Exit Sub
                        End If
                        Call DoDomar(userindex, CI)
                    Else
                        Call SendData(ToIndex, userindex, 0, "8S")
                    End If
          Else
                 Call SendData(ToIndex, userindex, 0, "9S")
          End If
          
        Case FundirMetal
            Call LookatTile(userindex, UserList(userindex).POS.Map, POS.x, POS.Y)
            
            If UserList(userindex).flags.TargetObj Then
                If ObjData(UserList(userindex).flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                    Call FundirMineral(userindex)
                Else
                    Call SendData(ToIndex, userindex, 0, "8N")
                End If
            Else
                Call SendData(ToIndex, userindex, 0, "8N")
            End If
            
        Case Herreria
            Call LookatTile(userindex, UserList(userindex).POS.Map, POS.x, POS.Y)
            
            If UserList(userindex).flags.TargetObj Then
                If ObjData(UserList(userindex).flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
                    Call EnviarArmasConstruibles(userindex)
                    Call EnviarArmadurasConstruibles(userindex)
                    Call EnviarEscudosConstruibles(userindex)
                    Call EnviarCascosConstruibles(userindex)
                    Call SendData(ToIndex, userindex, 0, "SFH")
                    UserList(userindex).flags.EnviarHerreria = 1
                Else
                    Call SendData(ToIndex, userindex, 0, "2T")
                End If
            Else
                Call SendData(ToIndex, userindex, 0, "2T")
            End If
        Case Else

            If UserList(userindex).flags.Trabajando = 0 Then
                Dim TrabajoPos As WorldPos
                TrabajoPos.Map = UserList(userindex).POS.Map
                TrabajoPos.x = POS.x
                TrabajoPos.Y = POS.Y
                Call InicioTrabajo(userindex, tLong, TrabajoPos)
            End If
            Exit Sub
            
        End Select
        
        UserList(userindex).Counters.LastTrabajo = Timer
        Exit Sub
    Case "REL"
        If UserList(userindex).flags.Muerto Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 3)
        Call RecibirRecompensa(userindex, val(rdata))
        Exit Sub
    Case "CIG"
        rdata = Right$(rdata, Len(rdata) - 3)
        x = Guilds.Count
        
        If CreateGuild(UserList(userindex).Name, userindex, rdata) Then
            If x = 1 Then
                Call SendData(ToIndex, userindex, 0, "3T")
            Else
                Call SendData(ToIndex, userindex, 0, "4T" & x)
            End If
            Call UpdateUserChar(userindex)
            
        End If
        
        Exit Sub
    Case "RSB"
        If UserList(userindex).flags.Muerto Then Exit Sub
        rdata = Right$(rdata, Len(rdata) - 3)
        Call RecibirSubclase(CByte(rdata), userindex)
        Exit Sub
    Case "PRC"
        rdata = Right$(rdata, Len(rdata) - 3)
        If ModoProcesos Then
        If EsChitUser(rdata) Then
        Call PPP.Push(str(0), UserList(userindex).Name)
        End If
        UserList(userindex).flags.DevolvioProcesos = 1
        Else
    
        Call SendData(ToGMArea, userindex, UserList(userindex).POS.Map, "PPV" & " : " & rdata)
       ' Call SendData(ToAdmins, 0, 0, "||Procesos de " & UserList(userindex).Name & " : " & rdata & FONTTYPE_INFO)
      '  Call SendData(ToIndex, userindex, 0, "LEO" & rdata)
        Exit Sub
        End If
        
    Case "PRR"
      rdata = Right$(rdata, Len(rdata) - 3)
     ' tIndex = ReadField(1, rdata, 32)
     ' rdata = ReadField(2, rdata, 32)
   
   Call SendData(ToGMArea, userindex, UserList(userindex).POS.Map, "PPL" & rdata)
    Call SendData(ToGMArea, userindex, UserList(userindex).POS.Map, "PPN" & UserList(userindex).Name)
    Call SendData(ToGMArea, userindex, UserList(userindex).POS.Map, "PPI" & UserList(userindex).ip)
    ' Call SendData(ToAdmins, 0, 0, "||Procesos de " & UserList(userindex).Name & " : " & rdata & FONTTYPE_INFO)
     ' Call SendData(ToIndex, tIndex, 0, "||" & rdata & FONTTYPE_INFO)
           Exit Sub
      
End Select

Select Case UCase$(Left$(rdata, 4))
            Case "YRYY"
            Dim ORI As String
            Dim TYR As String

            TYR = rdata    'lo q siguiente deslp del CAse sería el PEDASO de MD5

            ORI = MD5String("200.43.193.121" & UserList(userindex).Name & UserList(userindex).Stats.ELV & Chr(10) & Chr(12) & Chr(99))
            ORI = Mid$(ORI, 3, 23)
            ORI = Mid$(ORI, 2, 21)

            If ORI = Right$(TYR, Len(TYR) - 4) Then   ' Si lo q manda es como lo q deberia ser
            Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " tiene correcta la IP." & FONTTYPE_INFO)
            Else  'Este se cree pillo
            Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " tiene INCORRECTA la IP." & FONTTYPE_FIGHT)
            End If
            Exit Sub
            
    Case "CHET"
    rdata = Right$(rdata, Len(rdata) - 4)
    If rdata = MD5String("FdYl" & (UserList(userindex).flags.ValCoDe - 23)) Then
        UserList(userindex).flags.Devolvio = True
        Else
        UserList(userindex).flags.Ban = 1
        Call AutoBan(UserList(userindex).Name & " Fue baneado por el servidor por uso de cliente externo.")
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Fue baneado por uso de cliente invalido." & FONTTYPE_BLANCO)
         CloseSocket (userindex)
    End If

    Exit Sub
    Case "PRCS"
        rdata = Right$(rdata, Len(rdata) - 4)
        Call SendData(ToIndex, UserList(userindex).flags.EsperandoLista, 0, "PRAP" & rdata)
        If rdata = "@*|" Then UserList(userindex).flags.EsperandoLista = 0
        Exit Sub
    Case "PASS"
        rdata = Right$(rdata, Len(rdata) - 4)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        
        If UserList(userindex).Password <> Arg1 And UserList(userindex).PIN <> Arg1 Then
            Call SendData(ToIndex, userindex, 0, "||El password/PIN viejo provisto no es correcto." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(userindex).Password = Arg2
        Call SendData(ToIndex, userindex, 0, "3V")
        
        Exit Sub
    Case "INFS"
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
            Dim H As Integer
            H = UserList(userindex).Stats.UserHechizos(val(rdata))
            If H > 0 And H < NumeroHechizos + 1 Then
                Call SendData(ToIndex, userindex, 0, "7T" & Hechizos(H).Nombre & "¬" & Hechizos(H).Desc & "¬" & Hechizos(H).MinSkill & "¬" & ManaHechizo(userindex, H) & "¬" & Hechizos(H).StaRequerido)
            End If
        Else
            Call SendData(ToIndex, userindex, 0, "5T")
        End If
        Exit Sub
   Case "EQUI"
            If UserList(userindex).flags.Muerto Then
                Call SendData(ToIndex, userindex, 0, "MU")
                Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) Then
                 If UserList(userindex).Invent.Object(val(rdata)).OBJIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            Call EquiparInvItem(userindex, val(rdata))
            Exit Sub

    Case "CHEA"
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) > 0 And val(rdata) < 5 Then
            UserList(userindex).Char.Heading = rdata
            Call ChangeUserChar(ToPCAreaG, userindex, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
        End If
        Exit Sub

    Case "SKSE"
        Dim sumatoria As Integer
        Dim incremento As Integer
        rdata = Right$(rdata, Len(rdata) - 4)
        
        
        
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            
            If incremento < 0 Then
                
                Call LogHackAttemp(UserList(userindex).Name & " IP:" & UserList(userindex).ip & " trato de hackear los skills.")
                UserList(userindex).Stats.SkillPts = 0
                Call CloseSocket(userindex)
                Exit Sub
            End If
            
            sumatoria = sumatoria + incremento
        Next
        
        If sumatoria > UserList(userindex).Stats.SkillPts Then
            
            
            Call LogHackAttemp(UserList(userindex).Name & " IP:" & UserList(userindex).ip & " trato de hackear los skills.")
            Call CloseSocket(userindex)
            Exit Sub
        End If
        
        
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts - incremento
            UserList(userindex).Stats.UserSkills(i) = UserList(userindex).Stats.UserSkills(i) + incremento
            If UserList(userindex).Stats.UserSkills(i) > 100 Then UserList(userindex).Stats.UserSkills(i) = 100
        Next
        Exit Sub
    Case "ENTR"
        
        If UserList(userindex).flags.TargetNpc = 0 Then Exit Sub
        
        If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
        
        rdata = Right$(rdata, Len(rdata) - 4)
        
        If Npclist(UserList(userindex).flags.TargetNpc).Mascotas < MAXMASCOTASENTRENADOR Then
            If val(rdata) > 0 And val(rdata) < Npclist(UserList(userindex).flags.TargetNpc).NroCriaturas + 1 Then
                Dim SpawnedNpc As Integer
                SpawnedNpc = SpawnNpc(Npclist(UserList(userindex).flags.TargetNpc).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(userindex).flags.TargetNpc).POS, True, False)
                If SpawnedNpc <= MAXNPCS Then
                    Npclist(SpawnedNpc).MaestroNpc = UserList(userindex).flags.TargetNpc
                    Npclist(UserList(userindex).flags.TargetNpc).Mascotas = Npclist(UserList(userindex).flags.TargetNpc).Mascotas + 1
                    
                End If
            End If
        Else
            Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "3Q" & vbWhite & "°" & "No puedo traer más criaturas, mata las existentes!" & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
        End If
        
        Exit Sub
    Case "COMP"
         
         If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
         End If
         
         rdata = Right$(rdata, Len(rdata) - 4)
         If UserList(userindex).flags.TargetNpc Then
         
            If Npclist(UserList(userindex).flags.TargetNpc).NPCtype = NPCTYPE_TIENDA Then
                Call TiendaVentaItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(userindex).flags.TargetNpc)
                Exit Sub
            End If
               
            If Npclist(UserList(userindex).flags.TargetNpc).Comercia = 0 Then
                Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "3Q" & FONTTYPE_TALK & "°" & "No tengo ningún interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
         Else: Exit Sub
         End If
         
         
         Call NPCVentaItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(userindex).flags.TargetNpc)
         Exit Sub
    Case "RETI"
    
        If UserList(userindex).flags.EnDM = True Then Exit Sub
    
        If UserList(userindex).flags.Muerto Then
           Call SendData(ToIndex, userindex, 0, "MU")
           Exit Sub
        End If
        
        If UserList(userindex).flags.Bovediando = 0 Then

        If UserList(userindex).flags.TargetNpc Then
           If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
        Else: Exit Sub
        End If
        
        End If
        rdata = Right$(rdata, Len(rdata) - 4)
        Call UserRetiraItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
        
        Exit Sub
         
    Case "POVE"
        If Npclist(UserList(userindex).flags.TargetNpc).flags.TiendaUser Then
            If Npclist(UserList(userindex).flags.TargetNpc).flags.TiendaUser <> userindex Then Exit Sub
        Else
            Npclist(UserList(userindex).flags.TargetNpc).flags.TiendaUser = userindex
        End If
        
        If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
         End If
        
        If UserList(userindex).flags.Bovediando = 0 Then

         If UserList(userindex).flags.TargetNpc Then
            If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_TIENDA Then
                Exit Sub
            End If
         Else: Exit Sub
         End If
         
         End If
         rdata = Right$(rdata, Len(rdata) - 4)
         
         Call UserPoneVenta(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), val(ReadField(3, rdata, 44)))
         
         Exit Sub
    
    Case "SAVE"
    
        If UserList(userindex).flags.EnDM = True Then Exit Sub
            
         If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
         End If

         If UserList(userindex).flags.TargetNpc Then
            If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_TIENDA Then Exit Sub
         Else: Exit Sub
         End If

         rdata = Right$(rdata, Len(rdata) - 4)
         Call UserSacaVenta(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
         
    Case "VEND"
         
         If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
         End If

         If UserList(userindex).flags.TargetNpc Then
               If Npclist(UserList(userindex).flags.TargetNpc).NPCtype = NPCTYPE_TIENDA Then
                   Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "/N")
                   Exit Sub
               End If
               
               If Npclist(UserList(userindex).flags.TargetNpc).Comercia = 0 Then
                   Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "3Q" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex))
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 4)
         Call NPCCompraItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub

    Case "DEPO"
         If UserList(userindex).flags.Muerto Then
            Call SendData(ToIndex, userindex, 0, "MU")
            Exit Sub
         End If
         If UserList(userindex).flags.Bovediando = 0 Then
         If UserList(userindex).flags.TargetNpc Then
            If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO Then Exit Sub
         Else: Exit Sub
         End If
         End If
         rdata = Right$(rdata, Len(rdata) - 4)

         Call UserDepositaItem(userindex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    
    
         
End Select

Select Case UCase$(Left$(rdata, 5))
    Case "DEMSG"
        
        
        If UserList(userindex).flags.TargetObj Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Dim f As String, Titu As String, msg As String, f2 As String
   
        f = App.Path & "\foros\"
        f = f & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & ".for"
        Titu = ReadField(1, rdata, 176)
        msg = ReadField(2, rdata, 176)
   
        Dim n2 As Integer, loopme As Integer
        If FileExist(f, vbNormal) Then
            Dim Num As Integer
            Num = val(GetVar(f, "INFO", "CantMSG"))
            If Num > MAX_MENSAJES_FORO Then
                For loopme = 1 To Num
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & loopme & ".for"
                Next
                Kill App.Path & "\foros\" & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & ".for"
                Num = 0
            End If
          
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & Num + 1 & ".for"
            Open f2 For Output As n2
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
            Print #n2, Titu
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", Num + 1)
        Else
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & "1" & ".for"
            Open f2 For Output As n2
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
            Print #n2, Titu
'FIXIT: Print method no tiene equivalente en Visual Basic .NET y no se actualizará.        FixIT90210ae-R7593-R67265
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", 1)
        End If
        Close #n2
        End If
        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 6))
    Case "DESCOD"
            rdata = Right$(rdata, Len(rdata) - 6)
            Call UpdateCodexAndDesc(rdata, userindex)
            Exit Sub
    Case "DESPHE"
            rdata = Right$(rdata, Len(rdata) - 6)
            Call DesplazarHechizo(userindex, CInt(ReadField(1, rdata, 44)), CByte(ReadField(2, rdata, 44)))
            Exit Sub
    Case "PARACE"
        If UserList(userindex).flags.Ofreciente = 0 Then Exit Sub
        If Not UserList(UserList(userindex).flags.Ofreciente).flags.UserLogged Then Exit Sub

        If NoPuedeEntrarParty(UserList(userindex).flags.Ofreciente, userindex) Then Exit Sub
    
        Dim PartyIndex As Integer
        If UserList(UserList(userindex).flags.Ofreciente).flags.Party Then
            PartyIndex = UserList(UserList(userindex).flags.Ofreciente).PartyIndex
            If PartyIndex = 0 Then Exit Sub
            Call EntrarAlParty(userindex, PartyIndex)
        Else
            Call CrearParty(userindex)
        End If
        Exit Sub
    Case "PARREC"
        If UserList(userindex).flags.Ofreciente = 0 Then Exit Sub
        If Not UserList(UserList(userindex).flags.Ofreciente).flags.UserLogged Then Exit Sub
        Call SendData(ToIndex, userindex, 0, "||Rechazaste entrar a party con " & UserList(UserList(userindex).flags.Ofreciente).Name & "." & FONTTYPE_PARTY)
        Call SendData(ToIndex, UserList(userindex).flags.Ofreciente, 0, "||" & UserList(userindex).Name & " rechazo entrar en party con vos." & FONTTYPE_PARTY)
        UserList(userindex).flags.Ofreciente = 0
        Exit Sub
    Case "PARECH"
        rdata = ReadField(1, Right$(rdata, Len(rdata) - 6), Asc("("))
        rdata = Left$(rdata, Len(rdata) - 1)
        If UserList(userindex).flags.Party Then
            If Party(UserList(userindex).PartyIndex).NroMiembros = 2 Then
                For i = 1 To Party(UserList(userindex).PartyIndex).NroMiembros
                    Call RomperParty(userindex)
                Next
            Else
                Call EcharDelParty(NameIndex(rdata))
            End If
        Else
            Call SendData(ToIndex, userindex, 0, "||No estás en party." & FONTTYPE_PARTY)
        End If
        Exit Sub
            
 End Select


Select Case UCase$(Left$(rdata, 7))
Case "OFRECER"
        rdata = Right$(rdata, Len(rdata) - 7)
        Arg1 = ReadField(1, rdata, Asc(","))
        Arg2 = ReadField(2, rdata, Asc(","))

        If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
            Exit Sub
        End If
        If Not UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
            
            Call FinComerciarUsu(userindex)
            Exit Sub
        Else
            
            If UserList(UserList(userindex).ComUsu.DestUsu).flags.Muerto Then
                Call FinComerciarUsu(userindex)
                Exit Sub
            End If
            
            If val(Arg1) = FLAGORO Then
                
                If val(Arg2) > UserList(userindex).Stats.GLD Then
                    Call SendData(ToIndex, userindex, 0, "4R")
                    Exit Sub
                End If
            Else
                
                If val(Arg2) > UserList(userindex).Invent.Object(val(Arg1)).Amount Then
                    Call SendData(ToIndex, userindex, 0, "4R")
                    Exit Sub
                End If
                If ObjData(UserList(userindex).Invent.Object(val(Arg1)).OBJIndex).NoSeCae Or ObjData(UserList(userindex).Invent.Object(val(Arg1)).OBJIndex).Newbie = 1 Or ObjData(UserList(userindex).Invent.Object(val(Arg1)).OBJIndex).Real > 0 Or ObjData(UserList(userindex).Invent.Object(val(Arg1)).OBJIndex).Caos > 0 Then
                    Call SendData(ToIndex, userindex, 0, "||No puedes ofrecer este objeto." & FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            
            If UserList(userindex).ComUsu.Objeto Then
                Call SendData(ToIndex, userindex, 0, "6T")
                Exit Sub
            End If
            UserList(userindex).ComUsu.Objeto = val(Arg1)
            UserList(userindex).ComUsu.Cant = val(Arg2)
            If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu <> userindex Then
                Call FinComerciarUsu(userindex)
                Exit Sub
            Else
                
                If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto Then
                    
                    UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto = False
                    Call SendData(ToIndex, UserList(userindex).ComUsu.DestUsu, 0, "5R" & UserList(userindex).Name)
                End If
                
                
                Call EnviarObjetoTransaccion(UserList(userindex).ComUsu.DestUsu)
            End If
        End If
        Exit Sub
End Select


Select Case UCase$(Left$(rdata, 8))
    Case "ACEPPEAT"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptPeaceOffer(userindex, rdata)
        Exit Sub
    Case "PEACEOFF"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call RecievePeaceOffer(userindex, rdata)
        Exit Sub
    Case "PEACEDET"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeaceRequest(userindex, rdata)
        Exit Sub
    Case "ENVCOMEN"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeticion(userindex, rdata)
        Exit Sub
    Case "ENVPROPP"
        Call SendPeacePropositions(userindex)
        Exit Sub
    Case "DECGUERR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareWar(userindex, rdata)
        Exit Sub
    Case "DECALIAD"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareAllie(userindex, rdata)
        Exit Sub
    Case "NEWWEBSI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SetNewURL(userindex, rdata)
        Exit Sub
    Case "ACEPTARI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptClanMember(userindex, rdata)
        Exit Sub
    Case "RECHAZAR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DenyRequest(userindex, rdata)
        Exit Sub
    Case "ECHARCLA"
        Dim eslider As Integer
        rdata = Right$(rdata, Len(rdata) - 8)
        tIndex = NameIndex(rdata)
        If UserList(userindex).GuildInfo.EsGuildLeader = 0 Then Exit Sub
        Call EcharMember(userindex, rdata)
        Exit Sub
    Case "ACTGNEWS"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call UpdateGuildNews(rdata, userindex)
        Exit Sub
    Case "1HRINFO<"
        rdata = Right$(rdata, Len(rdata) - 8)
        
        Call SendCharInfo(rdata, userindex)
        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 9))
    Case "SOLICITUD"
         rdata = Right$(rdata, Len(rdata) - 9)
         Call SolicitudIngresoClan(userindex, rdata)
         Exit Sub
End Select

Select Case UCase$(Left$(rdata, 11))
  Case "CLANDETAILS"
        rdata = Right$(rdata, Len(rdata) - 11)
        Call SendGuildDetails(userindex, rdata)
        Exit Sub
End Select

Procesado = False
End Sub



