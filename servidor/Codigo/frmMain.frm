VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servidor FuriusAO - "
   ClientHeight    =   3540
   ClientLeft      =   1950
   ClientTop       =   1695
   ClientWidth     =   6615
   ControlBox      =   0   'False
   FillColor       =   &H80000004&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000007&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer TimerAntiChit 
      Interval        =   1000
      Left            =   3360
      Top             =   720
   End
   Begin VB.Timer TimerMeditar 
      Interval        =   400
      Left            =   2880
      Top             =   720
   End
   Begin VB.Data ADODB 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox i 
      Height          =   3180
      ItemData        =   "frmMain.frx":1042
      Left            =   5160
      List            =   "frmMain.frx":1049
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Mensaje BroadCast >>"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.Timer TimerRetos 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4200
         Top             =   600
      End
      Begin VB.Timer TimerSilencio 
         Interval        =   1000
         Left            =   3720
         Top             =   600
      End
      Begin VB.Timer TimerTrabaja 
         Interval        =   10000
         Left            =   4200
         Top             =   120
      End
      Begin VB.Timer CmdExec 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3240
         Tag             =   "S"
         Top             =   120
      End
      Begin VB.Timer UserTimer 
         Interval        =   1000
         Left            =   2760
         Top             =   120
      End
      Begin VB.Timer TimerFatuo 
         Interval        =   2500
         Left            =   3720
         Top             =   120
      End
      Begin VB.Timer tRevisarCabs 
         Left            =   10000
         Top             =   480
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   1800
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         Protocol        =   4
         RemoteHost      =   "fenixao.localstrike.com.ar"
         URL             =   "http://fenixao.localstrike.com.ar/descargas/Clave.txt"
         Document        =   "/descargas/Clave.txt"
         RequestTimeout  =   30
      End
      Begin VB.Label CantUsuarios 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblCantUsers 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Usuarios Online:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2400
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mensaje BroadCast:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
      Begin VB.TextBox BroadMsg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar Mensaje BroadCast"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   4695
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   6480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      X1              =   0
      X2              =   5160
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "&FuriusAo"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "SysTray Servidor"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar Servidor"
      End
      Begin VB.Menu mnuSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Cerrar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'fúriusao 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@furiusao.com.ar
'www.furiusao.com.ar
Option Explicit
Dim TimerChit As Integer
Public TLimpiarMapas As Integer
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer
Dim TimerS As Integer


Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Private Sub CmdExec_Timer()
On Error Resume Next

'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 Then
        If Not UserList(i).CommandsBuffer.IsEmpty Then Call HandleData(i, UserList(i).CommandsBuffer.Pop)
    End If
Next i

'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If

End Sub
Private Sub cmdMore_Click()

If cmdMore.Caption = "Mensaje BroadCast >>" Then
    Me.Height = 4395
    cmdMore.Caption = "<< Ocultar"
Else
    Me.Height = 2070
    cmdMore.Caption = "Mensaje BroadCast >>"
End If

End Sub

Private Sub Command1_Click()
Call SendData(ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub
Public Sub InitMain(f As Byte)

If f Then
    Call mnuSystray_Click
Else: frmMain.Show
End If

End Sub

Private Sub Fogatas_Timer()

End Sub

Private Sub Form_Load()

Call mnuSystray_Click
Codifico = RandomNumber(1, 99)
Torneo.MAXPARTICIPANTES = 12
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hwnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
'FIXIT: App.ThreadID property no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
'FIXIT: PopupMenu method no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
                PopupMenu mnuPopUp, , , , mnuMostrar
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub
Private Sub QuitarIconoSystray()
On Error Resume Next


Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray
'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 1 Then
    Call LimpiaWsApi(frmMain.hwnd)
'FIXIT: '#Else' no se actualiza de forma fiable a Visual Basic .NET                        FixIT90210ae-R2789-H1984
#Else
    Socket1.Cleanup
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If

Call DescargaNpcsDat

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next


Call LogMain(" Server cerrado")
End

End Sub

Private Sub mnuAyuda_Click()

End Sub

Private Sub mnuCerrar_Click()

Call SaveGuildsNew

If MsgBox("Si cierra el servidor puede provocar la perdida de datos." & vbCrLf & vbCrLf & "¿Desea hacerlo de todas maneras?", vbYesNo + vbExclamation, "Advertencia") = vbYes Then Call ApagarSistema

End Sub
Private Sub mnusalir_Click()

Call mnuCerrar_Click

End Sub
Public Sub mnuMostrar_Click()
On Error Resume Next

WindowState = vbNormal
Form_MouseMove 0, 0, 7725, 0

End Sub
Private Sub mnuServidor_Click()

frmServidor.Visible = True

End Sub
Private Sub mnuSystray_Click()
Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "Furius AO"
nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub
Private Sub Socket1_Blocking(Status As Integer, Cancel As Integer)
Cancel = True
End Sub
Private Sub Socket2_Connect(Index As Integer)

Set UserList(Index).CommandsBuffer = New CColaArray

End Sub
Private Sub Socket2_Disconnect(Index As Integer)

If UserList(Index).flags.UserLogged And _
    UserList(Index).Counters.Saliendo = False Then
    Call Cerrar_Usuario(Index)
Else: Call CloseSocket(Index)
End If

End Sub
Private Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)


'FIXIT: '#If' no se actualiza de forma fiable a Visual Basic .NET                          FixIT90210ae-R2789-H1984
#If UsarQueSocket = 0 Then
On Error GoTo ErrorHandler
Dim LoopC As Integer
Dim RD As String
'FIXIT: En Visual Basic .NET no se admiten matrices con límite inferior distinto de cero     FixIT90210ae-R9815-H1984
Dim rBuffer(1 To COMMAND_BUFFER_SIZE) As String
Dim CR As Integer
Dim tChar As String
Dim sChar As Integer
Dim eChar As Integer
Dim aux$
Dim OrigCad As String
Dim LenRD As Long

Call Socket2(Index).Read(RD, DataLength)

OrigCad = RD
LenRD = Len(RD)

If LenRD = 0 Then
    UserList(Index).AntiCuelgue = UserList(Index).AntiCuelgue + 1
    If UserList(Index).AntiCuelgue >= 150 Then
        UserList(Index).AntiCuelgue = 0
        Call LogError("!!!! Detectado bucle infinito de eventos socket2_read. cerrando indice " & Index)
        Socket2(Index).Disconnect
        Call CloseSocket(Index)
        Exit Sub
    End If
Else
    UserList(Index).AntiCuelgue = 0
End If

If Len(UserList(Index).RDBuffer) > 0 Then
    RD = UserList(Index).RDBuffer & RD
    UserList(Index).RDBuffer = ""
End If

sChar = 1
For LoopC = 1 To LenRD

    tChar = Mid$(RD, LoopC, 1)

    If tChar = ENDC Then
        CR = CR + 1
        eChar = LoopC - sChar
        rBuffer(CR) = Mid$(RD, sChar, eChar)
        sChar = LoopC + 1
    End If
        
Next LoopC

If Len(RD) - (sChar - 1) <> 0 Then UserList(Index).RDBuffer = Mid$(RD, sChar, Len(RD))

For LoopC = 1 To CR
    If ClientsCommandsQueue = 1 Then
        If Len(rBuffer(LoopC)) > 0 Then If Not UserList(Index).CommandsBuffer.Push(rBuffer(LoopC)) Then Call Cerrar_Usuario(Index)
    Else
        If UserList(Index).ConnID <> -1 Then
          Call HandleData(Index, rBuffer(LoopC))
        Else
          Exit Sub
        End If
    End If
Next LoopC

Exit Sub

ErrorHandler:
    Call LogError("Error en Socket read. " & Err.Description & " Numero paquetes:" & UserList(Index).NumeroPaquetesPorMiliSec & " . Rdata:" & OrigCad)
    Call CloseSocket(Index)
'FIXIT: '#End If' no se actualiza de forma fiable a Visual Basic .NET                      FixIT90210ae-R2789-H1984
#End If
End Sub

Private Sub TimerAntiChit_Timer()
TimerChit = TimerChit - 1

Dim XPos As WorldPos
XPos.Map = 200
XPos.X = 50
XPos.Y = 50
If TiempoMomia > 0 Then
    If TiempoMomia = 1 Then
    Call SpawnNpc(654, XPos, True, False)
    End If
    TiempoMomia = TiempoMomia - 1
End If


If TimerChit <= 0 Then
    If FPSBajos = True Then
    Call SendData(ToAdmins, 0, 0, "||Admins>> Puede que haya gente con FPS Menores a 5 /CheckFPS" & FONTTYPE_BLANCO)
    FPSBajos = False
    End If

CheckearDevolucion
TimerChit = 320
End If

If TLimpiarMapas <> 100 Then
TLimpiarMapas = TLimpiarMapas - 1
End If

If TLimpiarMapas = 60 Then
Call SendData(ToAll, 0, 0, "||60 segundos para la limpieza del mundo..." & FONTTYPE_furius)
End If

If TLimpiarMapas = 30 Then
Call SendData(ToAll, 0, 0, "||30 segundos para la limpieza del mundo..." & FONTTYPE_furius)
End If

If TLimpiarMapas = 15 Then
Call SendData(ToAll, 0, 0, "||15 segundos para la limpieza del mundo..." & FONTTYPE_furius)
End If

If TLimpiarMapas <= 0 Then
Call LimpiarMapas
TLimpiarMapas = 100
'TLimpiarMapas = 900
End If

End Sub

Private Sub TimerFatuo_Timer()
On Error GoTo Error
Dim i As Integer

For i = 1 To LastNPC
    If Npclist(i).flags.NPCActive And Npclist(i).Numero = 89 Then Npclist(i).CanAttack = 1
Next

Exit Sub

Error:
    Call LogError("Error en TimerFatuo: " & Err.Description)
End Sub
Private Sub TimerMeditar_Timer()
Dim i As Integer

For i = 1 To LastUser
    If UserList(i).flags.Meditando Then Call TimerMedita(i)
Next

End Sub
Sub TimerMedita(userindex As Integer)
Dim Cant As Single

If TiempoTranscurrido(UserList(userindex).Counters.tInicioMeditar) >= TIEMPO_INICIOMEDITAR Then
    Cant = UserList(userindex).Counters.ManaAcumulado + UserList(userindex).Stats.MaxMAN * (1 + UserList(userindex).Stats.UserSkills(Meditar) * 0.05) / 100
    If Cant <= 0.75 Then
        UserList(userindex).Counters.ManaAcumulado = Cant
        Exit Sub
    Else
        Cant = Round(Cant)
        UserList(userindex).Counters.ManaAcumulado = 0
    End If
    Call AddtoVar(UserList(userindex).Stats.MinMAN, Cant, UserList(userindex).Stats.MaxMAN)
    Call SendData(ToIndex, userindex, 0, "MN" & Cant)
    Call SubirSkill(userindex, Meditar)
    If UserList(userindex).Stats.MinMAN >= UserList(userindex).Stats.MaxMAN Then
        Call SendData(ToIndex, userindex, 0, "D9")
        Call SendData(ToIndex, userindex, 0, "MEDOK")
        UserList(userindex).flags.Meditando = False
        UserList(userindex).Char.FX = 0
        UserList(userindex).Char.loops = 0
        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
    End If
End If

Call SendUserMANA(userindex)

End Sub

Private Sub TimerRetos_Timer()
On Error Resume Next
TimerUVU = TimerUVU - 1
If TimerUVU < 1 Then
    If RetoEnCurso Then
        If UVUname > 0 Then
        Call WarpUserChar(UVUname, 160, 50, 50, True)
        End If
    RetoEnCurso = False
    End If
TimerRetos.Enabled = False
End If
End Sub

Private Sub TimerSilencio_Timer()
On Error GoTo ErrXD
TiempoReto = TiempoReto - 1
If TiempoReto < 1 Then
    If RetoEnCurso Then
        If UserList(RetoJ(1)).flags.EnReto And UserList(RetoJ(2)).flags.EnReto Then
            Call WarpUserChar(RetoJ(1), 160, 51, 50, True)
            Call WarpUserChar(RetoJ(2), 160, 50, 50, True)
            UserList(RetoJ(1)).flags.EnReto = 0
            UserList(RetoJ(2)).flags.EnReto = 0
            UserList(RetoJ(1)).flags.RetadoPor = 0
            UserList(RetoJ(2)).flags.RetadoPor = 0
            UserList(RetoJ(1)).flags.Retado = 0
            UserList(RetoJ(2)).flags.Retado = 0
            RetoEnCurso = False
            'TiempoReto = 0
        End If
    Else
        TiempoReto = 0
    End If

End If


    If TimerS <= 0 Then
        TimerS = 60
            Dim j As Integer
                For j = 1 To LastUser
                    If UserList(j).flags.Silenciado > 0 Then
                    UserList(j).flags.Silenciado = UserList(j).flags.Silenciado - 1
                    If UserList(j).flags.Silenciado = 0 Then
                       Call SendData(ToIndex, j, 0, "||Has sido liberado de tu silencio" & FONTTYPE_BLANCO)
                    End If
                End If
            Next j
        Exit Sub
    End If
    TimerS = TimerS - 1

Exit Sub
ErrXD:
Call LogError("Error en TimerSilencio: " & Err.Description)
End Sub

Private Sub TimerTrabaja_Timer()
Dim i As Integer
On Error GoTo Error

For i = 1 To LastUser
    If UserList(i).flags.Trabajando Then
        UserList(i).Counters.IdleCount = Timer
        
        Select Case UserList(i).flags.Trabajando
            Case Pesca
                Call DoPescar(i)
                    
            Case Talar
                Call DoTalar(i, ObjData(MapData(UserList(i).POS.Map, UserList(i).TrabajoPos.X, UserList(i).TrabajoPos.Y).OBJInfo.OBJIndex).ArbolElfico = 1)
    
            Case Mineria
                Call DoMineria(i, ObjData(MapData(UserList(i).POS.Map, UserList(i).TrabajoPos.X, UserList(i).TrabajoPos.Y).OBJInfo.OBJIndex).MineralIndex)
        End Select
    End If
Next
Exit Sub
Error:
    Call LogError("Error en TimerTrabaja: " & Err.Description)
    
End Sub
Private Sub UserTimer_Timer()
On Error Resume Next 'or GoTo Error
Static Andaban As Boolean, Contador As Single
Dim Andan As Boolean, UI As Integer, i As Integer

If CuentaRegresiva > 0 Then
    CuentaRegresiva = CuentaRegresiva - 1
    
    If CuentaRegresiva = 0 Then
        Call SendData(ToMap, 0, GMCuenta, "||YA!!!" & FONTTYPE_FIGHT)
        Me.Enabled = False
    Else
        Call SendData(ToMap, 0, GMCuenta, "||" & CuentaRegresiva & "..." & FONTTYPE_INFO)
    End If
End If

'For i = 1 To LastUser
'    If UserList(i).ConnID <> -1 Then DayStats.Segundos = DayStats.Segundos + 1
'Next

'If TiempoTranscurrido(Contador) >= 10 Then
'    Contador = Timer
'    Andan = EstadisticasWeb.EstadisticasAndando()
'    If Not Andaban And Andan Then Call InicializaEstadisticas
'    Andaban = Andan
'End If

For UI = 1 To LastUser
    If UserList(UI).flags.UserLogged And UserList(UI).ConnID <> -1 Then
       
        If UserList(UI).flags.Portal = 1 Then
'FIXIT: Declare 'mapachoto' and 'TPx' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
         Dim mapachoto, TPx, TPy As Integer
         mapachoto = UserList(UI).flags.PortalM
         TPx = UserList(UI).flags.PortalX
         TPy = UserList(UI).flags.PortalY
        If MapData(mapachoto, TPx, TPy).TileExit.Map > 0 Then
        Call EraseObj(ToMap, 0, mapachoto, MapData(mapachoto, TPx, TPy).OBJInfo.Amount, val(mapachoto), val(TPx), TPy)
        Call EraseObj(ToMap, 0, MapData(mapachoto, TPx, TPy).TileExit.Map, 1, MapData(mapachoto, TPx, TPy).TileExit.Map, MapData(mapachoto, TPx, TPy).TileExit.X, MapData(mapachoto, TPx, TPy).TileExit.Y)
        MapData(mapachoto, TPx, TPy).TileExit.Map = 0
        MapData(mapachoto, TPx, TPy).TileExit.X = 0
        MapData(mapachoto, TPx, TPy).TileExit.Y = 0
        End If
         
         UserList(UI).flags.Portal = 0
         UserList(UI).flags.PortalM = 0
         UserList(UI).flags.PortalY = 0
         UserList(UI).flags.PortalX = 0
        End If
         
         
         If UserList(UI).flags.Portal = 4 Then
         
'CREAMOS EL PORTAL
Dim Mapaf, Xf, Yf As Integer
Dim ET As Obj
ET.Amount = 1
ET.OBJIndex = 862 'portal luminoso
Mapaf = 1 'Hechizos(uh).TeleportXMap
Xf = 50 'Hechizos(uh).TeleportXX
Yf = 50 'hechizos(uh).TeleportXY
Call EraseObj(ToMap, UI, UserList(UI).flags.PortalM, 10000, UserList(UI).flags.PortalM, UserList(UI).flags.PortalX, UserList(UI).flags.PortalY)
Call MakeObj(ToMap, 0, UserList(UI).flags.PortalM, ET, UserList(UI).flags.PortalM, UserList(UI).flags.PortalX, UserList(UI).flags.PortalY)
MapData(UserList(UI).flags.PortalM, UserList(UI).flags.PortalX, UserList(UI).flags.PortalY).TileExit.Map = Mapaf
MapData(UserList(UI).flags.PortalM, UserList(UI).flags.PortalX, UserList(UI).flags.PortalY).TileExit.X = Xf
MapData(UserList(UI).flags.PortalM, UserList(UI).flags.PortalX, UserList(UI).flags.PortalY).TileExit.Y = Yf
Call SendData(ToPCArea, UI, UserList(UI).flags.PortalM, "TW" & SND_WARP)
'//// ACA TERMINAMOS DE CREARLO



         End If
         
         
         If UserList(UI).flags.Portal > 1 Then UserList(UI).flags.Portal = UserList(UI).flags.Portal - 1
         
         
        
        Call TimerPiquete(UI)
        Call TimerBoveda(UI)
        'ANTES DE ACA
        If UserList(UI).flags.Protegido > 1 Then Call TimerProtEntro(UI)
        If UserList(UI).flags.Encarcelado Then Call TimerCarcel(UI)
        If UserList(UI).flags.Muerto = 0 Then
            If UserList(UI).flags.Paralizado Then Call TimerParalisis(UI)
            If UserList(UI).flags.BonusFlecha Then Call TimerFlecha(UI)
            If UserList(UI).flags.Ceguera = 1 Then Call TimerCeguera(UI)
            If UserList(UI).flags.Envenenado = 1 Then Call TimerVeneno(UI)
            If UserList(UI).flags.Envenenado = 2 Then Call TimerVenenoDoble(UI)
            If UserList(UI).flags.Estupidez = 1 Then Call TimerEstupidez(UI)
            If UserList(UI).flags.AdminInvisible = 0 And UserList(UI).flags.Invisible = 1 And UserList(UI).flags.Oculto = 0 Then Call TimerInvisibilidad(UI)
            If UserList(UI).flags.Desnudo = 1 Then Call TimerFrio(UI)
            If UserList(UI).flags.TomoPocion Then Call TimerPocion(UI)
            If UserList(UI).flags.Transformado Then Call TimerTransformado(UI)
            If UserList(UI).NroMascotas Then Call TimerInvocacion(UI)
            If UserList(UI).flags.Oculto Then Call TimerOculto(UI)
            'UINVI
            If UserList(UI).Counters.uInvi > 0 Then UserList(UI).Counters.uInvi = UserList(UI).Counters.uInvi - 1
            'ULTI INVI
            
            Call TimerHyS(UI)
            Call TimerSanar(UI)
            Call TimerStamina(UI)
        End If
        If EnviarEstats Then
            Call SendUserStatsBox(UI)
            EnviarEstats = False
        End If
      
     
        Call TimerIdleCount(UI)
        
        If UserList(UI).Counters.Saliendo Then Call TimerSalir(UI)

    End If
Next

Exit Sub

Error:
    Call LogError("Error en UserTimer:" & Err.Description & " " & UI)
    
End Sub
Public Sub TimerOculto(userindex As Integer)
Dim ClaseBuena As Boolean

ClaseBuena = UserList(userindex).Clase = GUERRERO Or UserList(userindex).Clase = ARQUERO Or UserList(userindex).Clase = CAZADOR

If RandomNumber(1, 10 + UserList(userindex).Stats.UserSkills(Ocultarse) / 4 + 15 * Buleano(ClaseBuena) + 25 * Buleano(ClaseBuena And Not UserList(userindex).Clase = GUERRERO And UserList(userindex).Invent.ArmourEqpObjIndex = 360)) <= 5 Then
    UserList(userindex).flags.Oculto = 0
    UserList(userindex).flags.Invisible = 0
    Call SendData(ToMap, 0, UserList(userindex).POS.Map, ("V3" & UserList(userindex).Char.CharIndex & ",0"))
    Call SendData(ToIndex, userindex, 0, "V5")
End If

End Sub
Public Sub TimerStamina(userindex As Integer)

If UserList(userindex).Stats.MinSta < UserList(userindex).Stats.MaxSta And UserList(userindex).flags.Hambre = 0 And UserList(userindex).flags.Sed = 0 And UserList(userindex).flags.Desnudo = 0 Then
   If (Not UserList(userindex).flags.Descansar And TiempoTranscurrido(UserList(userindex).Counters.STACounter) >= StaminaIntervaloSinDescansar) Or _
   (UserList(userindex).flags.Descansar And TiempoTranscurrido(UserList(userindex).Counters.STACounter) >= StaminaIntervaloDescansar) Then
        UserList(userindex).Counters.STACounter = Timer
        UserList(userindex).Stats.MinSta = Minimo(UserList(userindex).Stats.MinSta + CInt(RandomNumber(5, Porcentaje(UserList(userindex).Stats.MaxSta, 15))), UserList(userindex).Stats.MaxSta)
        If TiempoTranscurrido(UserList(userindex).Counters.CartelStamina) >= 10 Then
            UserList(userindex).Counters.CartelStamina = Timer
            Call SendData(ToIndex, userindex, 0, "MV")
        End If
        EnviarEstats = True
    End If
End If

End Sub
Sub TimerTransformado(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Transformado) >= IntervaloInvisible Then
    Call DoTransformar(userindex)
End If

End Sub
Sub TimerInvisibilidad(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Invisibilidad) >= IntervaloInvisible Then
    Call SendData(ToIndex, userindex, 0, "V6")
    Call QuitarInvisible(userindex)
End If

End Sub
Sub TimerFlecha(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.BonusFlecha) >= 45 Then
    UserList(userindex).Counters.BonusFlecha = 0
    UserList(userindex).flags.BonusFlecha = False
    Call SendData(ToIndex, userindex, 0, "||Se acabó el efecto del Arco Encantado." & FONTTYPE_INFO)
End If

End Sub
Sub TimerPiquete(userindex As Integer)
On Error GoTo Errorent
If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).trigger = 5 Then
    UserList(userindex).Counters.PiqueteC = UserList(userindex).Counters.PiqueteC + 1
    If val(UserList(userindex).Counters.PiqueteC) Mod 5 = 0 Then Call SendData(ToIndex, userindex, 0, "9N")
    If val(UserList(userindex).Counters.PiqueteC) >= 25 Then
        UserList(userindex).Counters.PiqueteC = 0
        Call Encarcelar(userindex, 3)
    End If
Else: UserList(userindex).Counters.PiqueteC = 0
End If

Exit Sub
Errorent:
'Call LogError("Error en TimerPiquete : " & Err.Description)
End Sub
Sub TimerBoveda(userindex As Integer)
On Error GoTo Errorent
If UserList(userindex).flags.Muerto = 1 Then Exit Sub
If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).trigger = 8 Then
    Call SendData(ToIndex, userindex, 0, "+")
    UserList(userindex).Counters.tBoveda = UserList(userindex).Counters.tBoveda + 1
    If val(UserList(userindex).Counters.tBoveda) = 3 Then
        UserList(userindex).Counters.tBoveda = 4
        'Call IniciarDeposito(userindex, 3)
        Call SendData(ToIndex, userindex, 0, "SHWBP")
        UserList(userindex).flags.Bovediando = 1
        'UserList(userindex).flags.TargetNpc = NPCTYPE_BANQUERO
    End If
Else: UserList(userindex).Counters.tBoveda = 0: UserList(userindex).flags.Bovediando = 0
End If

Exit Sub
Errorent:
End Sub






Public Sub TimerProtEntro(userindex As Integer)
On Error GoTo Error

UserList(userindex).Counters.Protegido = UserList(userindex).Counters.Protegido - 1
If UserList(userindex).Counters.Protegido <= 0 Then UserList(userindex).flags.Protegido = 0

Exit Sub

Error:
    Call LogError("Error en TimerProtEntro" & " " & Err.Description)
End Sub
Sub TimerParalisis(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Paralisis) >= IntervaloParalizadoUsuario Then
    UserList(userindex).Counters.Paralisis = 0
    UserList(userindex).flags.Paralizado = 0
    Call SendData(ToIndex, userindex, 0, "P8")
End If

End Sub
Sub TimerCeguera(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Ceguera) >= IntervaloParalizadoUsuario / 10 Then
    UserList(userindex).Counters.Ceguera = 0
    UserList(userindex).flags.Ceguera = 0
    Call SendData(ToIndex, userindex, 0, "NSEGUE")
End If

End Sub
Sub TimerEstupidez(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Estupidez) >= IntervaloParalizadoUsuario Then
    UserList(userindex).Counters.Estupidez = 0
    UserList(userindex).flags.Estupidez = 0
    Call SendData(ToIndex, userindex, 0, "NESTUP")
End If

End Sub
Sub TimerCarcel(userindex As Integer)
If UserList(userindex).Counters.PenaMinar > 0 Then Exit Sub
If TiempoTranscurrido(UserList(userindex).Counters.Pena) >= UserList(userindex).Counters.TiempoPena Then
    UserList(userindex).Counters.TiempoPena = 0
    UserList(userindex).flags.Encarcelado = 0
    UserList(userindex).Counters.Pena = 0
    If UserList(userindex).POS.Map = Prision.Map Then
        Call WarpUserChar(userindex, Libertad.Map, Libertad.X, Libertad.Y, True)
        Call SendData(ToIndex, userindex, 0, "4P")
        Call SaleCarcelPique(userindex)
    End If
End If

End Sub

Sub TimerVenenoDoble(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Veneno) >= 2 Then
    If TiempoTranscurrido(UserList(userindex).flags.EstasEnvenenado) >= 8 Then
        UserList(userindex).flags.Envenenado = 0
        UserList(userindex).flags.EstasEnvenenado = 0
        UserList(userindex).Counters.Veneno = 0
    Else
        Call SendData(ToIndex, userindex, 0, "1M")
        UserList(userindex).Counters.Veneno = Timer
        If Not UserList(userindex).flags.Quest Then
            UserList(userindex).Stats.MinHP = Maximo(0, UserList(userindex).Stats.MinHP - 25)
            If UserList(userindex).Stats.MinHP = 0 Then
                Call UserDie(userindex)
            Else: EnviarEstats = True
            End If
        End If
    End If
End If

End Sub

Sub TimerVeneno(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Veneno) >= IntervaloVeneno Then
    If TiempoTranscurrido(UserList(userindex).flags.EstasEnvenenado) >= IntervaloVeneno * 10 Then
        UserList(userindex).flags.Envenenado = 0
        UserList(userindex).flags.EstasEnvenenado = 0
        UserList(userindex).Counters.Veneno = 0
    Else
        Call SendData(ToIndex, userindex, 0, "1M")
        UserList(userindex).Counters.Veneno = Timer
        If Not UserList(userindex).flags.Quest Then
            UserList(userindex).Stats.MinHP = Maximo(0, UserList(userindex).Stats.MinHP - RandomNumber(1, 5))
            If UserList(userindex).Stats.MinHP = 0 Then
                Call UserDie(userindex)
            Else: EnviarEstats = True
            End If
        End If
    End If
End If

End Sub
Public Sub TimerFrio(userindex As Integer)

If UserList(userindex).flags.Privilegios > 1 Then Exit Sub

If TiempoTranscurrido(UserList(userindex).Counters.Frio) >= IntervaloFrio Then
    UserList(userindex).Counters.Frio = Timer
    If MapInfo(UserList(userindex).POS.Map).Terreno = Nieve Then
        If TiempoTranscurrido(UserList(userindex).Counters.CartelFrio) >= 5 Then
            UserList(userindex).Counters.CartelFrio = Timer
            Call SendData(ToIndex, userindex, 0, "1K")
        End If
        If Not UserList(userindex).flags.Quest Then
            UserList(userindex).Stats.MinHP = Maximo(0, UserList(userindex).Stats.MinHP - Porcentaje(UserList(userindex).Stats.MaxHP, 5))
            EnviarEstats = True
            If UserList(userindex).Stats.MinHP = 0 Then
                Call SendData(ToIndex, userindex, 0, "1L")
                Call UserDie(userindex)
            End If
        End If
    End If
    Call QuitarSta(userindex, Porcentaje(UserList(userindex).Stats.MaxSta, 5))
    If TiempoTranscurrido(UserList(userindex).Counters.CartelFrio) >= 10 Then
        UserList(userindex).Counters.CartelFrio = Timer
        Call SendData(ToIndex, userindex, 0, "FR")
    End If
    EnviarEstats = True
End If

End Sub
Sub TimerPocion(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).flags.DuracionEfecto) >= 45 Then
    UserList(userindex).flags.DuracionEfecto = 0
    UserList(userindex).flags.TomoPocion = False
    UserList(userindex).Stats.UserAtributos(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad)
    UserList(userindex).Stats.UserAtributos(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza)
    Call UpdateFuerzaYAg(userindex)
End If

End Sub
Public Sub TimerHyS(userindex As Integer)
Dim EnviaInfo As Boolean

If UserList(userindex).flags.Privilegios > 1 Or (UserList(userindex).Clase = TALADOR And UserList(userindex).Recompensas(1) = 2) Or UserList(userindex).flags.Quest Then Exit Sub

If TiempoTranscurrido(UserList(userindex).Counters.AGUACounter) >= IntervaloSed Then
    If UserList(userindex).flags.Sed = 0 Then
        UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU - 10
        If UserList(userindex).Stats.MinAGU <= 0 Then
            UserList(userindex).Stats.MinAGU = 0
            UserList(userindex).flags.Sed = 1
        End If
        EnviaInfo = True
    End If
    UserList(userindex).Counters.AGUACounter = Timer
End If

If TiempoTranscurrido(UserList(userindex).Counters.COMCounter) >= IntervaloHambre Then
    If UserList(userindex).flags.Hambre = 0 Then
        UserList(userindex).Counters.COMCounter = Timer
        UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinHam - 10
        If UserList(userindex).Stats.MinHam <= 0 Then
            UserList(userindex).Stats.MinHam = 0
            UserList(userindex).flags.Hambre = 1
        End If
        EnviaInfo = True
    End If
    UserList(userindex).Counters.COMCounter = Timer
End If

If EnviaInfo Then Call EnviarHambreYsed(userindex)

End Sub
Sub TimerSanar(userindex As Integer)

If (UserList(userindex).flags.Descansar And TiempoTranscurrido(UserList(userindex).Counters.HPCounter) >= SanaIntervaloDescansar) Or _
     (Not UserList(userindex).flags.Descansar And TiempoTranscurrido(UserList(userindex).Counters.HPCounter) >= SanaIntervaloSinDescansar) Then
    If (Not Lloviendo Or Not Intemperie(userindex)) And UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP And UserList(userindex).flags.Hambre = 0 And UserList(userindex).flags.Sed = 0 Then
        If UserList(userindex).flags.Descansar Then
            UserList(userindex).Stats.MinHP = Minimo(UserList(userindex).Stats.MaxHP, UserList(userindex).Stats.MinHP + Porcentaje(UserList(userindex).Stats.MaxHP, 20))
            If UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MinHP And UserList(userindex).Stats.MaxSta = UserList(userindex).Stats.MinSta Then
                Call SendData(ToIndex, userindex, 0, "DOK")
                Call SendData(ToIndex, userindex, 0, "DN")
                UserList(userindex).flags.Descansar = False
            End If
        Else
            UserList(userindex).Stats.MinHP = Minimo(UserList(userindex).Stats.MaxHP, UserList(userindex).Stats.MinHP + Porcentaje(UserList(userindex).Stats.MaxHP, 5))
        End If
        Call SendData(ToIndex, userindex, 0, "1N")
        EnviarEstats = True
    End If
    UserList(userindex).Counters.HPCounter = Timer
End If
    
End Sub
Sub TimerInvocacion(userindex As Integer)
Dim i As Integer
Dim NpcIndex As Integer

If UserList(userindex).flags.Privilegios > 0 Or UserList(userindex).flags.Quest Then Exit Sub

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(i) Then
        NpcIndex = UserList(userindex).MascotasIndex(i)
        If Npclist(NpcIndex).Contadores.TiempoExistencia > 0 And TiempoTranscurrido(Npclist(NpcIndex).Contadores.TiempoExistencia) >= IntervaloInvocacion + 10 * Buleano(Npclist(NpcIndex).Numero = 92) Then
        Call MuereNpc(NpcIndex, 0)
        End If
      End If
Next

End Sub
Public Sub TimerIdleCount(userindex As Integer)

If UserList(userindex).flags.Privilegios = 0 And UserList(userindex).flags.Trabajando = 0 And TiempoTranscurrido(UserList(userindex).Counters.IdleCount) >= IntervaloParaConexion And Not UserList(userindex).Counters.Saliendo Then
    Call SendData(ToIndex, userindex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
    Call SendData(ToIndex, userindex, 0, "FINOK")
    Call CloseSocket(userindex)
End If

End Sub
Sub TimerSalir(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Salir) >= IntervaloCerrarConexion Then
    Call SendData(ToIndex, userindex, 0, "FINOK")
    Call CloseSocket(userindex)
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub
