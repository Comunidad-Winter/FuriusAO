VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   4455
      Left            =   6360
      ScaleHeight     =   4395
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox NoticiaS 
         BackColor       =   &H80000006&
         ForeColor       =   &H80000005&
         Height          =   4365
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Text            =   "frmConnect.frx":1982
         Top             =   0
         Visible         =   0   'False
         Width           =   4365
      End
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   7545
      Width           =   2895
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Left            =   120
      MaxLength       =   20
      TabIndex        =   1
      Top             =   6555
      Width           =   2895
   End
   Begin VB.Image imgWeb 
      Height          =   1575
      Left            =   240
      MouseIcon       =   "frmConnect.frx":1998
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image imgGetPass 
      Height          =   375
      Left            =   9240
      MouseIcon       =   "frmConnect.frx":1CA2
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   315
      Index           =   0
      Left            =   9240
      MouseIcon       =   "frmConnect.frx":1FAC
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   315
      Index           =   1
      Left            =   3240
      MouseIcon       =   "frmConnect.frx":22B6
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   315
      Index           =   2
      Left            =   9240
      MouseIcon       =   "frmConnect.frx":25C0
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   1890
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Integer, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Private Sub Command1_Click()
'Password.Left = RandomNumber(1, 9150)
'Password.Top = RandomNumber(1, 7500)
'Password.Show
'Password.SetFocus

'End Sub
Private Sub Form_Activate()
txtUser.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    Call PlayWaveDS(SND_CLICK)
            
    If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    
    If frmConnect.MousePointer = 11 Then
    frmConnect.MousePointer = 1
        Exit Sub
    End If
    
    
    UserName = txtUser.Text
    Dim aux As String
    aux = txtPass.Text
    UserPassword = MD5String(aux)
    If CheckUserData(False) = True Then
        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        
        EstadoLogin = Normal
        Me.MousePointer = 11
        frmMain.Socket1.Connect
    End If
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'frmCliente.Visible = True'
'frmCliente.Show , Me
If KeyCode = 27 Then
    frmCargando.Show
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "Cerrando FuriusAO.", 255, 150, 50, 1, 0, 1
    
    Call SaveGameini
    frmConnect.MousePointer = 1
    frmMain.MousePointer = 1
    prgRun = False
    
    AddtoRichTextBox frmCargando.Status, "Liberando recursos..."
    frmCargando.Refresh
    LiberarObjetosDX
    AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, 0, 1
    AddtoRichTextBox frmCargando.Status, "¡¡Gracias por jugar FuriusAO!!", 255, 150, 50, 1, 0, 1
    frmCargando.Refresh
    Call UnloadAllForms
End If

End Sub

Private Sub Form_Load()
'NoticiaS.Text = FrmIntro.Ipxd.OpenURL("http://www.furiusao.com.ar/NOTICIAS.TXT")
    EngineRun = False
    
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next

 IntervaloPaso = 0.19
 IntervaloUsar = 0.14
 Picture = LoadPicture(DirGraficos & "conectar.jpg")
  



End Sub

Private Sub Image1_Click(Index As Integer)

CurServer = 0
Unload Password

'Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0

        If Musica = 0 Then
            CurMidi = DirMidi & "7.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If

       
        EstadoLogin = dados
        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        Me.MousePointer = 11
        frmMain.Socket1.Connect
        
    Case 1
     '  If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
     'FuriusAO Generaba problema al desconectase y vovlerse a conectar. Lo dejamos con SocketReset
        frmMain.Socket1.Disconnect
        If frmConnect.MousePointer = 11 Then
        frmConnect.MousePointer = 1
            Exit Sub
        End If
        
        
        
        UserName = txtUser.Text
        Dim aux As String
        aux = txtPass.Text
        UserPassword = MD5String(aux)
        If CheckUserData(False) = True Then
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
            
            EstadoLogin = Normal
            Me.MousePointer = 11
            frmMain.Socket1.Connect
        End If
        
    Case 2
        Call ShellExecute(Me.Hwnd, "open", "http://www.furiusao.com.ar/paneluser/index.php", "", "", 1)

End Select

End Sub
Private Sub imgGetPass_Click()

Call ShellExecute(Me.Hwnd, "open", "http://www.furiusao.com.ar/paneluser/index.php", "", "", 1)

End Sub
Private Sub imgWeb_Click()

Call ShellExecute(Me.Hwnd, "open", "http://www.furiusao.com.ar", "", "", 1)

End Sub


Function RandomNumber(ByVal LowerBound As Single, ByVal UpperBound As Single) As Single
Randomize Timer
RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function




