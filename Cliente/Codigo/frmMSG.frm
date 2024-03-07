VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   0  'None
   Caption         =   "GM Messenger"
   ClientHeight    =   7230
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox mensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1575
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4680
      Width           =   5295
   End
   Begin VB.TextBox GM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Text            =   "Cualquier GM disponible"
      Top             =   2470
      Width           =   2775
   End
   Begin VB.ComboBox categoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00111720&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmMSG.frx":0000
      Left            =   2880
      List            =   "frmMSG.frx":0013
      TabIndex        =   1
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmMSG.frx":0065
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmMSG.frx":036F
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   2175
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "SGM.gif")

End Sub
Private Sub Image1_Click()
Dim GMs As String

If categoria.ListIndex = -1 Then
    MsgBox "El motivo del mensaje no es válido"
    Exit Sub
End If

If Len(mensaje.Text) > 250 Then
    MsgBox "La longitud del mensaje debe tener menos de 250 carácteres."
    Exit Sub
End If

If Len(mensaje.Text) < 10 Then
    MsgBox "La longitud del mensaje es muy corta."
    Exit Sub
End If

If Len(GM.Text) = 0 Or GM.Text = "Cualquier GM disponible" Then
    GMs = "Ninguno"
Else: GMs = GM.Text
End If

If Len(mensaje.Text) = 0 Then
    MsgBox "Debes ingresar un mensaje."
    Exit Sub
End If

Call SendData("CS" & mensaje.Text)

If NoMandoElMsg = 0 Then
    mensaje.Text = ""
    GM.Text = "Cualquier GM disponible"
    categoria.List(categoria.ListIndex) = ""
    AddtoRichTextBox frmMain.RecTxt, "El mensaje fue enviado, Rogamos tengas paciencia y no escribas más de un mensaje sobre el mismo tema, tambien puedes usar el soporte web, en www.furiusao.com.ar", 252, 151, 53, 1, 0
    Unload Me
Else
    Call MsgBox("El mensaje es demasiado largo, por favor resumilo.")
End If

End Sub



Private Sub Label7_Click()

End Sub

Private Sub mensaje_Change()
mensaje.Text = LTrim(mensaje.Text)
End Sub


Private Sub mensaje_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 209) And (KeyAscii <> 241) And (KeyAscii <> 8) And (KeyAscii <> 32) And (KeyAscii <> 164) And (KeyAscii <> 165) Then
    If (Index <> 6) And ((KeyAscii < 40 Or KeyAscii > 122) Or (KeyAscii > 90 And KeyAscii < 96)) Then
        KeyAscii = 0
    End If
End If

 KeyAscii = Asc((Chr(KeyAscii)))
End Sub
