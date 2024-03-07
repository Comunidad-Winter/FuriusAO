Attribute VB_Name = "Anti_Cheat"
Option Explicit

Dim Usando_cheat As Long
Public Mando_cheat(10) As String
Public Procesos(150) As String

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Integer = 260

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
szExeFile As String * MAX_PATH
End Type

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Function IscheatRunning(cheattype) As Boolean
   IscheatRunning = (FindWindow(vbNullString, cheattype) <> 0)
End Function

Function verify_cheats2()
Dim i As Double
Dim Num As Double

Usando_cheat = "0"

For i = 1 To 400
Num = Num + 0.01
If IscheatRunning("!xSpeed.net +" & Left(CStr(i), 1) & "." & Right(CStr(Num), 2)) = True Then
Usando_cheat = "1"
send_cheats2
End If
Next

If IscheatRunning("makro-piringulete") = True Then
Usando_cheat = "1"
send_cheats2
End If

If IscheatRunning("makro K33") = True Then
Usando_cheat = "1"
send_cheats2
End If

If IscheatRunning("makro-Piringulete 2003") = True Then
Usando_cheat = "1"
send_cheats2
End If

If IscheatRunning("macrocrack <gonza_vi@hotmail.com>") = True Then
Usando_cheat = "1"
send_cheats2
End If

If IscheatRunning("windows speeder") = True Then
Usando_cheat = "2"
send_cheats2
End If

If IscheatRunning("Speeder - Unregistered") = True Then
Usando_cheat = "2"
send_cheats2
End If

If IscheatRunning("A Speeder") = True Then
Usando_cheat = "2"
send_cheats2
End If

If IscheatRunning("speeder") = True Then
Usando_cheat = "3"
send_cheats2
End If

If IscheatRunning("argentum-pesca 0.2b por manchess") = True Then
Usando_cheat = "4"
send_cheats2
End If

If IscheatRunning("speeder XP - softwrap version") = True Then
Usando_cheat = "5"
send_cheats2
End If

If IscheatRunning("aoflechas") = True Then
Usando_cheat = "6"
send_cheats2
End If

If IscheatRunning("Macro") = True Then
Usando_cheat = "6"
send_cheats2
End If

If IscheatRunning("Macro.Pete") = True Then
Usando_cheat = "6"
send_cheats2
End If

If IscheatRunning("Macro Pete") = True Then
Usando_cheat = "6"
send_cheats2
End If

If IscheatRunning("Macro 2005") = True Then
Usando_cheat = "7"
send_cheats2
End If

If IscheatRunning("Cheat Engine V5.1.1") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.2") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V5.0") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V4.4") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V4.4 German Add-On") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V4.3") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V4.2") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V4.1.1") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V3.3") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V3.2") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine V3.1") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Samples Macros - EZ Macros") = True Then
Usando_cheat = "6"
send_cheats2
End If

If IscheatRunning("MacroMaker") = True Then
Usando_cheat = "6"
send_cheats2
End If

If IscheatRunning("solocovo?") = True Then
Usando_cheat = "1"
send_cheats2
End If

If IscheatRunning("korven") = True Then
Usando_cheat = "1"
send_cheats2
End If

If IscheatRunning("kizSada") = True Then
Usando_cheat = "1"
send_cheats2
End If

If IscheatRunning("By ^[cavallero]^") = True Then
Usando_cheat = "1"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.0") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("0RK4M VERSION 1.5") = True Then
Usando_cheat = "6"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.1.1") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.1") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.1.0") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.1.2") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.3") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.3.1") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.3.2") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.3.3") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.3.4") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.3.5") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine 5.4") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Cheat Engine") = True Then
Usando_cheat = "8"
send_cheats2
End If

If IscheatRunning("Serbio Engine") = True Then
Usando_cheat = "8"
send_cheats2
End If
End Function

Function verify_cheats()
Dim i As Double
Dim Num As Double

Usando_cheat = "0"

For i = 1 To 400
Num = Num + 0.01
If IscheatRunning("!xSpeed.net +" & Left(CStr(i), 1) & "." & Right(CStr(Num), 2)) = True Then
Usando_cheat = "1"
send_cheats
End If
Next

If IscheatRunning("makro-piringulete") = True Then
Usando_cheat = "1"
send_cheats
End If

If IscheatRunning("makro K33") = True Then
Usando_cheat = "1"
send_cheats
End If

If IscheatRunning("makro-Piringulete 2003") = True Then
Usando_cheat = "1"
send_cheats
End If

If IscheatRunning("macrocrack <gonza_vi@hotmail.com>") = True Then
Usando_cheat = "1"
send_cheats
End If

If IscheatRunning("windows speeder") = True Then
Usando_cheat = "2"
send_cheats
End If

If IscheatRunning("MacroMaker") = True Then
Usando_cheat = "6"
send_cheats
End If

If IscheatRunning("Speeder - Unregistered") = True Then
Usando_cheat = "2"
send_cheats
End If

If IscheatRunning("A Speeder") = True Then
Usando_cheat = "2"
send_cheats
End If

If IscheatRunning("speeder") = True Then
Usando_cheat = "3"
send_cheats
End If

If IscheatRunning("argentum-pesca 0.2b por manchess") = True Then
Usando_cheat = "4"
send_cheats
End If

If IscheatRunning("speeder XP - softwrap version") = True Then
Usando_cheat = "5"
send_cheats
End If

If IscheatRunning("aoflechas") = True Then
Usando_cheat = "6"
send_cheats
End If

If IscheatRunning("Macro") = True Then
Usando_cheat = "6"
send_cheats
End If

If IscheatRunning("Macro 2005") = True Then
Usando_cheat = "7"
send_cheats
End If

If IscheatRunning("Cheat Engine V5.1.1") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V5.0") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V4.4") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V4.4 German Add-On") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V4.3") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V4.2") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V4.1.1") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V3.3") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V3.2") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine V3.1") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Samples Macros - EZ Macros") = True Then
Usando_cheat = "6"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.0") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("0RK4M VERSION 1.5") = True Then
Usando_cheat = "6"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.1.1") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.1.0") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.1.2") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.2") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.3") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.3.1") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.3.2") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.3.3") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.3.4") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.3.5") = True Then
Usando_cheat = "8"
send_cheats
End If

If IscheatRunning("Cheat Engine 5.4") = True Then
Usando_cheat = "8"
send_cheats
End If

End Function

Function send_cheats()

'If (Mando_cheat(Usando_cheat)) = False Then

Mando_cheat(Usando_cheat) = True
SendData ("@" & Usando_cheat)
MsgBox "Programa externo detectado. Argentum Online se cerrará.", vbCritical, "Atención!"
UnloadAllForms
'End If
End Function

Function send_cheats2()

'If (Mando_cheat(Usando_cheat)) = False Then

Mando_cheat(Usando_cheat) = True
'SendData ("@" & Usando_cheat)
MsgBox "Programa externo detectado. Argentum Online se cerrará.", vbCritical, "Atención!"
End

'End If

End Function

Sub ListApps()
Dim a As Integer, i As Integer, lista As String
         Dim hSnapShot As Long
         Dim uProceso As PROCESSENTRY32
         Dim R As Long

         hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
         If hSnapShot = 0 Then Exit Sub
         uProceso.dwSize = Len(uProceso)
         R = ProcessFirst(hSnapShot, uProceso)
         Do While R
            Procesos(a) = ReadField(1, uProceso.szExeFile, Asc("."))
            If UCase$(Procesos(a)) = "!XSPEEDNET.EXE" Or _
            UCase$(Procesos(a)) = "!XSPEEDNET" Or _
            UCase$(Procesos(a)) = "CHEAT ENGINE" Then
            Usando_cheat = "2"
            Call send_cheats
            End If
            a = a + 1
            R = ProcessNext(hSnapShot, uProceso)
         Loop

         For i = 2 To UBound(Procesos)
         If Procesos(i) <> "" Then
         lista = lista & Procesos(i) & ","
         End If
         Next
         SendData "€" & UCase$(lista)

         Call CloseHandle(hSnapShot)
End Sub

Sub ListApps2()
Dim a As Integer, i As Integer, lista As String
         Dim hSnapShot As Long
         Dim uProceso As PROCESSENTRY32
         Dim R As Long

         hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
         If hSnapShot = 0 Then Exit Sub
         uProceso.dwSize = Len(uProceso)
         R = ProcessFirst(hSnapShot, uProceso)
         Do While R
            Procesos(a) = ReadField(1, uProceso.szExeFile, Asc("."))
            If UCase$(Procesos(a)) = "!XSPEEDNET.EXE" Or _
            UCase$(Procesos(a)) = "!XSPEEDNET" Or _
            UCase$(Procesos(a)) = "CHEAT ENGINE" Then
            Usando_cheat = "2"
            send_cheats2 (Usando_cheat)
            End If
            a = a + 1
            R = ProcessNext(hSnapShot, uProceso)
         Loop

         For i = 2 To UBound(Procesos)
         If Procesos(i) <> "" Then
         lista = lista & Procesos(i) & ","
         End If
         Next
         'SendData "€" & UCase$(lista)

         Call CloseHandle(hSnapShot)
End Sub

Public Function Encryptar(Texto As String, val As Long) As String
Dim i As Integer
Dim sec As String

For i = 1 To Len(Texto)
    sec = Asc(Mid(Texto, i, 1)) + val
    sec = ChrW(sec)
    Mid(Texto, i, 1) = sec
Next

Encryptar = Texto

End Function

Public Function Desencryptar(Texto As String, val As Long) As String
Dim i As Integer
Dim sec As String

For i = 1 To Len(Texto)
    sec = Asc(Mid(Texto, i, 1)) - val
    sec = ChrW(sec)
    Mid(Texto, i, 1) = sec
Next

Desencryptar = Texto

End Function

    

'Option Explicit
'
'Dim Usando_cheat As Long
'Public Mando_cheat(5) As String
'Public Procesos(50) As String
'
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'
'Public Const TH32CS_SNAPPROCESS As Long = 2&
'Public Const MAX_PATH As Integer = 260
'
'Public Type PROCESSENTRY32
'dwSize As Long
'cntUsage As Long
'th32ProcessID As Long
'th32DefaultHeapID As Long
'th32ModuleID As Long
'cntThreads As Long
'th32ParentProcessID As Long
'pcPriClassBase As Long
'dwFlags As Long
'szExeFile As String * MAX_PATH
'End Type
'
'Public Declare Function CreateToolhelpSnapshot Lib "Kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
'Public Declare Function ProcessFirst Lib "Kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
'Public Declare Function ProcessNext Lib "Kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
'Public Declare Sub CloseHandle Lib "Kernel32" (ByVal hPass As Long)
'
'Public Function IscheatRunning(cheattype) As Boolean
'   IscheatRunning = (FindWindow(vbNullString, cheattype) <> 0)
'End Function
'
'Function verify_cheats()
'Usando_cheat = "0"
'
'If IscheatRunning("makro-piringulete") = True Then
'Usando_cheat = "1"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("makro K33") = True Then
'Usando_cheat = "1"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("makro-Piringulete 2003") = True Then
'Usando_cheat = "1"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("macrocrack <gonza_vi@hotmail.com>") = True Then
'Usando_cheat = "1"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("windows speeder") = True Then
'Usando_cheat = "2"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Speeder - Unregistered") = True Then
'Usando_cheat = "2"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("A Speeder") = True Then
'Usando_cheat = "2"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("speeder") = True Then
'Usando_cheat = "3"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("argentum-pesca 0.2b por manchess") = True Then
'Usando_cheat = "4"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("speeder XP - softwrap version") = True Then
'Usando_cheat = "5"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("aoflechas") = True Then
'Usando_cheat = "6"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Macro") = True Then
'Usando_cheat = "6"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Macro 2005") = True Then
'Usando_cheat = "7"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V5.1.1") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V5.0") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V4.4") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V4.4 German Add-On") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V4.3") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V4.2") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V4.1.1") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V3.3") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V3.2") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine V3.1") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Cheat Engine") = True Then
'Usando_cheat = "8"
'send_cheats (Usando_cheat)
'End If
'
'If IscheatRunning("Samples Macros - EZ Macros") = True Then
'Usando_cheat = "6"
'send_cheats (Usando_cheat)
'End If
'
'End Function
'
'Function send_cheats()
'
''If (Mando_cheat(Usando_cheat)) = False Then
'
'Mando_cheat(Usando_cheat) = True
'SendData ("@" & Usando_cheat)
'MsgBox "Programa externo detectado. Argentum Online se cerrará.", vbCritical, "Atención!"
'UnloadAllForms
''End If
'End Function
'
'Sub ListApps()
'Dim a As Integer, I As Integer, lista As String
'         Dim hSnapShot As Long
'         Dim uProceso As PROCESSENTRY32
'         Dim r As Long
'
'         hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
'         If hSnapShot = 0 Then Exit Sub
'         uProceso.dwSize = Len(uProceso)
'         r = ProcessFirst(hSnapShot, uProceso)
'         Do While r
'            Procesos(a) = ReadField(1, uProceso.szExeFile, Asc("."))
'            If UCase$(Procesos(a)) = "!XSPEEDNET.EXE" Or _
'            UCase$(Procesos(a)) = "!XSPEEDNET" Then
'            Usando_cheat = "2"
'            send_cheats (Usando_cheat)
'            End If
'            a = a + 1
'            r = ProcessNext(hSnapShot, uProceso)
'         Loop
'
'         For I = 2 To UBound(Procesos)
'         If Procesos(I) <> "" Then
'         lista = lista & Procesos(I) & ","
'         End If
'         Next
'         SendData "€" & UCase$(lista)
'
'         Call CloseHandle(hSnapShot)
'End Sub

Function Aocrypt(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String


    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    Aocrypt = strBuff
End Function

Function Aodecrypt(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    Aodecrypt = strBuff
End Function



