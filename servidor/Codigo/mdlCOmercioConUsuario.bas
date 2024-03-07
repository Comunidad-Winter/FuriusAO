Attribute VB_Name = "mdlCOmercioConUsuario"

Option Explicit

Public Type tCOmercioUsuario
    DestUsu As Integer
    DestNick As String
    Objeto As Integer
    Cant As Long
    
    Acepto As Boolean
End Type
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
On Error GoTo errhandler

If UserList(Origen).ComUsu.DestUsu = Destino And _
   UserList(Destino).ComUsu.DestUsu = Origen Then
    
    Call UpdateUserInv(True, Origen, 0)
    
    Call SendData(ToIndex, Origen, 0, "INITCOMUSU")
    UserList(Origen).flags.Comerciando = True

    
    Call UpdateUserInv(True, Destino, 0)
    
    Call SendData(ToIndex, Destino, 0, "INITCOMUSU")
    UserList(Destino).flags.Comerciando = True
Else
    
    Call SendData(ToIndex, Destino, 0, "||" & UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR." & FONTTYPE_TALK)
    UserList(Destino).flags.TargetUser = Origen
    
End If

Exit Sub
errhandler:

End Sub
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer)
Dim ObjInd As Integer
Dim ObjCant As Long

ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Cant
If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
    ObjInd = iORO
Else
    ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).OBJIndex
End If

If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub

If ObjInd > 0 And ObjCant Then
    Call SendData(ToIndex, AQuien, 0, "COMUSUINV" & 1 & "," & ObjInd & "," & ObjData(ObjInd).Name & "," & ObjCant & "," & 0 & "," & ObjData(ObjInd).GrhIndex & "," _
    & ObjData(ObjInd).ObjType & "," _
    & ObjData(ObjInd).MaxHIT & "," _
    & ObjData(ObjInd).MinHIT & "," _
    & ObjData(ObjInd).MaxDef & "," _
    & ObjData(ObjInd).Valor \ 3)
End If

End Sub
Public Sub FinComerciarUsu(userindex As Integer)

If userindex = 0 Then Exit Sub

With UserList(userindex)
    If .ComUsu.DestUsu Then Call SendData(ToIndex, userindex, 0, "FINCOMUSUOK")
    .ComUsu.Acepto = False
    .ComUsu.Cant = 0
    .ComUsu.DestUsu = 0
    .ComUsu.Objeto = 0
    .ComUsu.DestNick = ""
    .flags.Comerciando = False
End With

End Sub
Public Sub AceptarComercioUsu(userindex As Integer)
Dim Obj1 As Obj, Obj2 As Obj
Dim OtroUserIndex As Integer
Dim TerminarAhora As Boolean


   
OtroUserIndex = UserList(userindex).ComUsu.DestUsu

If OtroUserIndex <= 0 Then
    Call FinComerciarUsu(userindex)
    Exit Sub
End If

If UserList(userindex).flags.EnDM = True Then Exit Sub
If UserList(OtroUserIndex).flags.EnDM = True Then Exit Sub
     
     
TerminarAhora = (UserList(OtroUserIndex).ComUsu.DestUsu <> userindex) Or _
                (UserList(OtroUserIndex).Name <> UserList(userindex).ComUsu.DestNick) Or _
                (UserList(userindex).Name <> UserList(OtroUserIndex).ComUsu.DestNick) Or _
                (Not UserList(OtroUserIndex).flags.UserLogged Or Not UserList(userindex).flags.UserLogged)

If TerminarAhora Then
    Call FinComerciarUsu(userindex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If

UserList(userindex).ComUsu.Acepto = True

If Not UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto Then
    Call SendData(ToIndex, userindex, 0, "||El otro usuario aun no ha aceptado tu oferta." & FONTTYPE_TALK)
    Exit Sub
End If

If UserList(userindex).ComUsu.Objeto = FLAGORO Then
    Obj1.OBJIndex = iORO
    If UserList(userindex).ComUsu.Cant > UserList(userindex).Stats.GLD Then
        Call SendData(ToIndex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj1.Amount = UserList(userindex).ComUsu.Cant
    Obj1.OBJIndex = UserList(userindex).Invent.Object(UserList(userindex).ComUsu.Objeto).OBJIndex
    If Obj1.Amount > UserList(userindex).Invent.Object(UserList(userindex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If
If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    Obj2.OBJIndex = iORO
    If UserList(OtroUserIndex).ComUsu.Cant > UserList(OtroUserIndex).Stats.GLD Then
        Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
Else
    Obj2.Amount = UserList(OtroUserIndex).ComUsu.Cant
    Obj2.OBJIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).OBJIndex
    If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
        Call SendData(ToIndex, OtroUserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
        TerminarAhora = True
    End If
End If

If TerminarAhora Then
    Call FinComerciarUsu(userindex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
End If


If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
    
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Cant
    Call SendUserORO(OtroUserIndex)
    
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Cant
    Call SendUserORO(userindex)
Else
    
    If Not MeterItemEnInventario(userindex, Obj2) Then Call TirarItemAlPiso(UserList(userindex).POS, Obj2)
    Call QuitarObjetos(Obj2.OBJIndex, Obj2.Amount, OtroUserIndex)
End If


If UserList(userindex).ComUsu.Objeto = FLAGORO Then
    
    UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - UserList(userindex).ComUsu.Cant
    Call SendUserORO(userindex)
    
    UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(userindex).ComUsu.Cant
    Call SendUserORO(OtroUserIndex)
Else
    
    If Not MeterItemEnInventario(OtroUserIndex, Obj1) Then Call TirarItemAlPiso(UserList(OtroUserIndex).POS, Obj1)
    Call QuitarObjetos(Obj1.OBJIndex, Obj1.Amount, userindex)
End If



Call UpdateUserInv(True, userindex, 0)
Call UpdateUserInv(True, OtroUserIndex, 0)

Call FinComerciarUsu(userindex)
Call FinComerciarUsu(OtroUserIndex)
 
End Sub



