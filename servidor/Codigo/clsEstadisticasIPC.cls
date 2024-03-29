VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEstadisticasIPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'FIXIT: As Any no se admite en Visual Basic .NET. Utilice un tipo específico.              FixIT90210ae-R5608-H1984
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_USER = &H400

Private sMensaje As Long, hVentana As Long, hVentanaMia As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDnext = 2


Public Enum EstaNotificaciones
    CANTIDAD_ONLINE = 1
    RECORD_USUARIOS = 2
    UPTIME_SERVER = 3
    CANTIDAD_MAPAS = 4
    EVENTO_NUEVO_CLAN = 5
    
    HANDLE_WND_SERVER = 100
End Enum



Private Function BuscaVentana(Wnd As Long, str As String) As Long
Dim W As Long
Dim L As Long
Dim T As String


W = GetWindow(Wnd, GW_HWNDFIRST)

While W <> 0
   L = GetWindowTextLength(W)

   If L Then
      T = Space$(L + 1)
      L = GetWindowText(W, T, L + 1)

      If Left$(T, Len(str)) = str Then
         BuscaVentana = W
         Exit Function
      End If
   End If

   W = GetWindow(W, GW_HWNDnext)
Wend

BuscaVentana = 0

End Function

Public Function Informar(ByVal QueCosa As EstaNotificaciones, ByVal Parametro As Long) As Long
Call BuscaWndEstadisticas
Informar = SendMessageLong(hVentana, sMensaje, QueCosa, Parametro)

End Function

Public Function EstadisticasAndando() As Boolean
Dim Ret As Long

Call BuscaWndEstadisticas

EstadisticasAndando = (hVentana <> 0)

End Function

Public Sub Inicializa(ByVal hwnd As Long)
hVentanaMia = hwnd
sMensaje = RegisterWindowMessage("EstadisticasAO")

End Sub

Private Sub BuscaWndEstadisticas()
hVentana = BuscaVentana(hVentanaMia, "Servidor de estadisticas AO")

End Sub
