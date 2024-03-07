VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   7800
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox TransTecho 
      BackColor       =   &H80000007&
      Caption         =   "Check2"
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Configurar teclas"
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   6960
      Width           =   1455
   End
   Begin VB.OptionButton SBMP 
      BackColor       =   &H00000000&
      Caption         =   "Option2"
      Height          =   195
      Left            =   3240
      TabIndex        =   21
      Top             =   4260
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton SJPG 
      BackColor       =   &H80000007&
      Caption         =   "Option1"
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   4240
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Información"
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   17
      Top             =   4560
      Width           =   3255
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "¿Necesitás &ayuda?"
         Height          =   225
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&www.furiusao.com.ar"
         Height          =   255
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   2520
      Max             =   100
      TabIndex        =   16
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Silencio"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   2400
      Max             =   255
      Min             =   40
      TabIndex        =   11
      Top             =   6480
      Value           =   255
      Width           =   1455
   End
   Begin VB.PictureBox Clanes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2235
      MouseIcon       =   "frmOpciones.frx":1234D
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   3300
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   720
      MouseIcon       =   "frmOpciones.frx":12657
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   9
      Top             =   3675
      Width           =   335
   End
   Begin VB.PictureBox PictureSanado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2235
      MouseIcon       =   "frmOpciones.frx":12961
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   7
      Top             =   2880
      Width           =   335
   End
   Begin VB.PictureBox PictureRecuMana 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   720
      MouseIcon       =   "frmOpciones.frx":12C6B
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   2400
      Width           =   335
   End
   Begin VB.PictureBox PictureVestirse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   720
      MouseIcon       =   "frmOpciones.frx":12F75
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   2880
      Width           =   335
   End
   Begin VB.PictureBox PictureMenosCansado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2235
      MouseIcon       =   "frmOpciones.frx":1327F
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   2400
      Width           =   335
   End
   Begin VB.PictureBox PictureNoHayNada 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2235
      MouseIcon       =   "frmOpciones.frx":13589
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   1920
      Width           =   335
   End
   Begin VB.PictureBox PictureOcultarse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   720
      MouseIcon       =   "frmOpciones.frx":13893
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   1950
      Width           =   335
   End
   Begin VB.PictureBox PictureFxs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   2280
      MouseIcon       =   "frmOpciones.frx":13B9D
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   1200
      Width           =   335
   End
   Begin VB.PictureBox PictureMusica 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   335
      Left            =   720
      MouseIcon       =   "frmOpciones.frx":13EA7
      MousePointer    =   99  'Custom
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1200
      Width           =   335
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Efectos"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   1215
      Left            =   1560
      TabIndex        =   13
      Top             =   5760
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   2143
      _Version        =   393216
      BorderStyle     =   1
      Orientation     =   1
      Max             =   4000
      TickStyle       =   2
      TickFrequency   =   500
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   1215
      Left            =   840
      TabIndex        =   15
      Top             =   5760
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   2143
      _Version        =   393216
      BorderStyle     =   1
      Orientation     =   1
      LargeChange     =   500
      SmallChange     =   500
      Max             =   4000
      TickStyle       =   2
      TickFrequency   =   500
      TextPosition    =   1
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000C0C0&
      Height          =   375
      Left            =   600
      Top             =   4185
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C0C0&
      Height          =   2175
      Left            =   600
      Top             =   1905
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      Height          =   375
      Left            =   600
      Top             =   1185
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   600
      MouseIcon       =   "frmOpciones.frx":141B1
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmOpciones.frx":144BB
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   975
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Clanesx As Byte
' función Api para aplicar la transparencia a la ventana
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal Hwnd As Long, _
     ByVal crKey As Long, _
     ByVal bAlpha As Byte, _
     ByVal dwFlags As Long) As Long
' Funciones api para los estilos de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal Hwnd As Long, _
     ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal Hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
'constantes
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Public Function Transparencia(ByVal Hwnd As Long, Valor As Integer) As Long
On Local Error GoTo ErrSub
Dim Estilo As Long
    Estilo = GetWindowLong(Hwnd, GWL_EXSTYLE)
    Estilo = Estilo Or WS_EX_LAYERED
    SetWindowLong Hwnd, GWL_EXSTYLE, Estilo
    SetLayeredWindowAttributes Hwnd, 0, Valor, LWA_ALPHA
    Transparencia = 0
If Err Then
    Transparencia = 2
End If
Exit Function
ErrSub:
MsgBox Err.Description, vbCritical, "Error"
End Function

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Command1_Click()
frmCustomKeys.Show , frmMain
End Sub

Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
Me.Slider2.Value = 4000
Me.Slider2.Value = 4000
Musica = 1
FX = 1

Else
 If bLluvia(UserMap) = 0 Then
    If bRain Then
        IMC.Stop
        End If
    End If
Musica = 0
FX = 0
End If
End Sub

Private Sub Clanes_Click()
If Clanesx = 0 Then
    Clanesx = 1
    Clanes.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    Clanesx = 0
    Clanes.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub cmdAyuda_Click()
frmHlp.Show vbModeless, frmOpciones
End Sub

Private Sub cmdWeb_Click()
ShellExecute Me.Hwnd, "open", "http://www.furiusao.com.ar", "", "", 1
End Sub

Private Sub Form_Activate()

'tam.Caption = Val(frmMain.FontSize)
End Sub

Private Sub Form_Load()


Me.Picture = LoadPicture(DirGraficos & "Opciones.gif")

If Musica = 0 Then
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If FX = 0 Then
    PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If NoRes = 1 Then
    Picture1.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    Picture1.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelOcultarse = 1 Then
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelMenosCansado = 1 Then
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelVestirse = 1 Then
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelNoHayNada = 1 Then
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelRecuMana = 1 Then
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

If CartelSanado = 1 Then
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
If Clanes = 1 Then
    Clanes.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    Clanes.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
End Sub

Private Sub HScroll1_Change()
 Call Transparencia(FormConsola.Hwnd, HScroll1.Value)
 Call Transparencia(FormTalk.Hwnd, HScroll1.Value)
 Call Transparencia(FormListOpciones.Hwnd, HScroll1.Value)
 Call Transparencia(FormBarInv.Hwnd, HScroll1.Value)
 Call Transparencia(FormInv.Hwnd, HScroll1.Value)
 Call Transparencia(FormInfo.Hwnd, HScroll1.Value)

End Sub



Private Sub HScroll2_Change()
Call OpenMixer
Call actualizavolumen
End Sub

Private Sub Image1_Click()

Me.Visible = False

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Image3_Click()
On Error GoTo error
Dim tamaño As String
tamaño = InputBox("Indique el tamaño ( en numeros), por ejemplo, 10, sin comillas")
tam.Caption = Val(tamaño)
Prueba.FontSize = Val(tamaño)
error:
'MsgBox "Err al cambiar la letra"
End Sub

Private Sub Image4_Click()
If Prueba.FontBold = False Then
Prueba.FontBold = True
'pruebax.FontBold = True
Exit Sub
ElseIf Prueba.FontBold = True Then
Prueba.FontBold = False
'pruebax.FontBold = False
Exit Sub
End If
End Sub

Private Sub Image5_Click()
If Prueba.FontItalic = False Then
Prueba.FontItalic = True
'pruebax.FontItalic = True
Exit Sub
ElseIf Prueba.FontItalic = True Then
Prueba.FontItalic = False
'pruebax.FontItalic = True
Exit Sub
End If
End Sub

Private Sub Image6_Click()
frmMain.Font = Prueba.Font
frmMain.FontSize = Prueba.FontSize
frmMain.Font = Prueba.FontBold
frmMain.FontItalic = Prueba.FontItalic
End Sub

Private Sub Label11_Click()

ShellExecute Me.Hwnd, "open", "http://www.furiusao.com.ar/Manual/", "", "", 1

End Sub



Private Sub Label16_Click()

End Sub

Private Sub Picture1_Click()

If NoRes = 0 Then
    NoRes = 1
    Picture1.Picture = LoadPicture(DirGraficos & "tick1.gif")
    Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 1)
Else
    NoRes = 0
    Picture1.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Call WriteVar(App.Path & "/Init/Opciones.opc", "CONFIG", "ModoVentana", 0)
End If

MsgBox "Este cambio hará efecto recién la próxima vez que ejecutes el juego."

End Sub

Private Sub PictureFxs_Click()

Select Case FX
    Case 0
        FX = 1
        PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
    Case 1
        FX = 0
        PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
End Select

End Sub
Private Sub PictureMenosCansado_Click()

If CartelMenosCansado = 0 Then
    CartelMenosCansado = 1
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelMenosCansado = 0
    PictureMenosCansado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If

Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "MenosCansado", str(CartelMenosCansado))

End Sub

Private Sub PictureMusica_Click()

If Not IsPlayingCheck Then
    Musica = 0
    Play_Midi
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    Musica = 1
    Stop_Midi
    PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If




End Sub

Private Sub PictureNoHayNada_Click()
If CartelNoHayNada = 0 Then
    CartelNoHayNada = 1
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelNoHayNada = 0
    PictureNoHayNada.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "NoHayNada", str(CartelNoHayNada))

End Sub

Private Sub PictureOcultarse_Click()

If CartelOcultarse = 0 Then
    CartelOcultarse = 1
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelOcultarse = 0
    PictureOcultarse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Ocultarse", str(CartelOcultarse))
End Sub

Private Sub PictureRecuMana_Click()
If CartelRecuMana = 0 Then
    CartelRecuMana = 1
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelRecuMana = 0
    PictureRecuMana.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "RecuMana", str(CartelRecuMana))

End Sub

Private Sub PictureSanado_Click()
If CartelSanado = 0 Then
    CartelSanado = 1
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelSanado = 0
    PictureSanado.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Sanado", str(CartelSanado))

End Sub

Private Sub PictureVestirse_Click()
If CartelVestirse = 0 Then
    CartelVestirse = 1
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick1.gif")
Else
    CartelVestirse = 0
    PictureVestirse.Picture = LoadPicture(DirGraficos & "tick2.gif")
End If
Call WriteVar(App.Path & "/Init/Opciones.opc", "CARTELES", "Vestirse", str(CartelVestirse))

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton Then

      dX = X

      dy = Y

      bmoving = True

   End If

   

End Sub

 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving And ((X <> dX) Or (Y <> dy)) Then

      Move Left + (X - dX), Top + (Y - dy)

   End If

   

End Sub

 

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then

      bmoving = False

   End If

   

End Sub

Private Sub Slider1_Change()
If Me.Slider1.Value <= 0 Then
Me.Slider1.Text = "Normal"
Me.Check1.Value = 0
ElseIf Me.Slider1.Value > 0 And Me.Slider1.Value < 2000 Then
Me.Slider1.Text = "Bajo"
Me.Check1.Value = 0
ElseIf Me.Slider1.Value > 2000 And Me.Slider1.Value < 4000 Then
Me.Slider1.Text = "Muy Bajo"
Me.Check1.Value = 0
Else
Me.Slider1.Text = "Silencio"
End If

End Sub


Private Sub Slider2_Change()
If Me.Slider1.Value >= 0 Then
Me.Slider2.Text = "Normal"
Me.Check1.Value = 0
ElseIf Me.Slider1.Value < 0 And Me.Slider2.Value < 2000 Then
Me.Slider2.Text = "Bajo"
Me.Check1.Value = 0
ElseIf Me.Slider1.Value > 2000 And Me.Slider2.Value < 4000 Then
Me.Slider2.Text = "Muy Bajo"
Me.Check1.Value = 0
Else
Me.Slider2.Text = "Silencio"
End If

Perf.SetMasterVolume -Me.Slider2.Value
End Sub

