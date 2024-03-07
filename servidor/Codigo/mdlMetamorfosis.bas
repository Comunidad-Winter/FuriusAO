Attribute VB_Name = "modMetamorfosis"
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub DoMetamorfosis(userindex As Integer)

metacuerpo = RandomNumber(1, 10)

Select Case (metacuerpo)
    Case 1
        metacuerpo = 9
    Case 2
        metacuerpo = 11
    Case 3
        metacuerpo = 42
    Case 4
        metacuerpo = 243
    Case 5
        metacuerpo = 149
    Case 6
        metacuerpo = 151
    Case 7
        metacuerpo = 155
    Case 8
        metacuerpo = 157
    Case 9
        metacuerpo = 159
    Case 10
        metacuerpo = 141
End Select

UserList(userindex).flags.Transformado = 1
UserList(userindex).Counters.Transformado = Timer

Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, val(userindex), metacuerpo, 0, UserList(userindex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
        
If UserList(userindex).flags.Desnudo Then UserList(userindex).flags.Desnudo = 0

Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_MORPH)
Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXWARPMORPH & "," & 0)

End Sub
Sub DoTransformar(userindex As Integer, Optional ByVal FX As Boolean = True)

UserList(userindex).flags.Transformado = 0
UserList(userindex).Counters.Transformado = 0

If UserList(userindex).Invent.ArmourEqpObjIndex = 0 Then
    Call DarCuerpoDesnudo(userindex)
Else
    UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
End If

If UserList(userindex).Invent.CascoEqpObjIndex = 0 Then
   UserList(userindex).Char.CascoAnim = NingunCasco
Else
    UserList(userindex).Char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
End If

If UserList(userindex).Invent.EscudoEqpObjIndex = 0 Then
   UserList(userindex).Char.ShieldAnim = NingunEscudo
Else
    UserList(userindex).Char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
End If

If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
   UserList(userindex).Char.WeaponAnim = NingunArma
Else
    UserList(userindex).Char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
End If

Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)

If FX Then
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW" & SND_WARPMORPH)
    Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXWARPMORPH & "," & 0)
End If

End Sub
