VERSION 5.00
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form leo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de control "
   ClientHeight    =   5220
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog10 
      Left            =   3960
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   9240
      TabIndex        =   17
      Top             =   4800
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4800
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   -120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdupload 
      Caption         =   "Subir"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Foto"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
   Begin Captura.wndCaptura FotoPro 
      Left            =   4920
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IMAGEN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Label20"
      Height          =   375
      Left            =   10680
      TabIndex        =   13
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblRESPONSE 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12120
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label ip 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   4350
      Width           =   3015
   End
   Begin VB.Label nombre 
      Caption         =   "nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   4350
      Width           =   4695
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   4800
      Y1              =   -360
      Y2              =   4680
   End
   Begin VB.Label Label3 
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Nick 
      Caption         =   "Nick:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Aplicaciones 
      Caption         =   "Aplicaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Procesos y ruta de ejecucion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mnuPrc 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnucer 
         Caption         =   "Cerrar "
      End
      Begin VB.Menu mnuinf 
         Caption         =   "Informacion"
      End
      Begin VB.Menu mnutime 
         Caption         =   "Tiempo de ejecucion"
      End
   End
End
Attribute VB_Name = "leo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim net As Inet
'Dim flag As Boolean
'Dim State As Integer
'Dim AccessType As String


Private Sub Command1_Click()

Dim X As Integer
FotoPro.Area = Ventana
FotoPro.Captura
FotoPro.Area = Ventana
FotoPro.Captura


         
For X = 1 To 1000
If Not FileExist(App.Path & "\Procesos", vbDirectory) Then Call MkDir$(App.Path & "\Procesos")
If Not FileExist(App.Path & "/Procesos/" & leo.nombre.Caption & X & ".bmp", vbNormal) Then Exit For
Next
Call SavePicture(FotoPro.Imagen, App.Path & "/Procesos/" & leo.nombre.Caption & X & ".bmp")
Text2.Text = App.Path & "\Procesos\" & leo.nombre.Caption & X & ".bmp"
Call AddtoRichTextBox(frmMain.RecTxt, "La imagen de los procesos de de screenshots bajo el nombre de " & leo.nombre.Caption & X & ".bmp", 255, 150, 50, False, False, False)




End Sub



Private Sub Command2_Click()

End Sub

Private Sub CommonDialog10_Click()

End Sub

Private Sub Form_Load()

  'se abre el formulario y se autoconfigura INET1.
'  leo.Text3.Text = ""
'  Inet1.URL = "ftp://furiusao.com.ar" 'Direcion del FTP
'  Inet1.Proxy = "http://furiusao.com.ar" 'COmo se entra al FTP mediante web
'  Inet1.UserName = "furiusao" 'User
'  Inet1.Password = "leo159753" ' Password
'  Inet1.RequestTimeout = 60 'Tiempo para conexion OUT
'  Inet1.AccessType = icDirect 'Tipo de conexion de acceso
'  Inet1.Protocol = icFTP 'Protocolo del INET
'  Inet1.RemotePort = 21 'Puerto del FTP
  'Call Conectar
End Sub
'Private Sub Conectar()
'    On Error GoTo FTPError
'        Inet1.URL = "ftp://furiusao.com.ar"
'          Inet1.UserName = "furiusao"
'          Inet1.Password = "leo159753"
'          Inet1.Execute , "", "POST", ""
'FTPError:
'            Select Case Err.Number
'            Case 35754
'                 'MsgBox Err.Description
'                 Me.Label20.Caption = Err.Description
'          '  Case Default
'                 'MsgBox "Connected Successfully."
'                 Me.Label20.Caption = "Conexion exitosa, esperando imagenes....  "
'            End Select
'End Sub
  
Private Sub Label1_Click()
If MsgBox("¿Está seguro de cerrar el proceso seleccionado al usuario?", vbYesNo) = vbYes Then
Call SendData("/CERRARPROCESO " & nombre & "@" & ReadField(2, List2.List(List2.ListIndex), Asc(":")))
If List2.ListIndex <> -1 Then
   'Eliminamos el elemento que se encuentra seleccionado
   List2.RemoveItem List2.ListIndex
End If
End If
End Sub



'Private Sub cmdupload_Click()
'
' Dim localfile As String
'        Dim remoteFile As String
'        Dim FileName As String
'        Dim filepath, fileid As String
'        Dim seperatorPos, s1 As Integer
'  On Error Resume Next
'
 '          ChDir App.Path
 '       'Open Text1.Text For Input As #2
 '        localfile = Text2.Text
 '               seperatorPos = InStrRev(localfile, "\", -1, vbTextCompare)
 '               FileName = Right(localfile, Len(localfile) - seperatorPos)
 '               s1 = InStrRev(localfile, "\", -1, vbTextCompare)
 '               fileid = Right(localfile, Len(localfile) - s1)
 '               MsgBox fileid
 '               remoteFile = "" & FileName
 '               Dim str As String
 '               Inet1.Execute , "PUT """ & localfile & """ """ & remoteFile & """"
 '               Me.lblRESPONSE.Caption = " Esperando respuesta: " & Inet1.ResponseInfo & " Respuesta: " & Inet1.ResponseCode
 '               Me.lblRESPONSE.Caption = "Archivo Subido correctamente.."
 '               Text3.Text = "[IMG]" & "http://ao.localstrike.com.ar/procesos/" & remoteFile & "[IMG]"
              
'End Sub


'Private Sub Inet1_StateChanged(ByVal State As Integer)'

'    On Error Resume Next
'    Dim vtData As Variant
'    Select Case State
'           Case icNone
'           Case icResolvingHost: Me.lblRESPONSE.Caption = "Resolviendo host"
'           Case icHostResolved: Me.lblRESPONSE.Caption = "Host encontrado"
'           Case icConnecting: Me.lblRESPONSE.Caption = "Conectando..."
'           Case icConnected: Me.lblRESPONSE.Caption = "Conectado"
'           Case icResponseReceived: Me.lblRESPONSE.Caption = "Enviando Archivo..."
'           Case icDisconnecting: Me.lblRESPONSE.Caption = "Desconectando..."
'           Case icDisconnected: Me.lblRESPONSE.Caption = "Desconectado"
'           Case icError: MsgBox "Error:" & Inet1.ResponseCode & " " & Inet1.ResponseInfo
'           Case icResponseCompleted:  Me.lblRESPONSE.Caption = "Imagen subida correctamente...."
'     End Select
'
 '        Me.lblRESPONSE.Refresh
 '      '  Me.lblRESPONSE.Caption = ""
 '     Err.Clear
 ''
'End Sub

'Private Sub inetconnect()
'  Inet1.URL = "ftp://furiusao.com.ar"
'  Inet1.Proxy = "http://furiusao.com.ar"
'  Inet1.UserName = "furiusao"
'  Inet1.Password = "leo159753"
'  Inet1.RequestTimeout = 60
'  Inet1.AccessType = icDirect
'  Inet1.Protocol = icFTP
'  Inet1.RemotePort = 21
'
'End Sub




