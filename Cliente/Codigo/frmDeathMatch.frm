VERSION 5.00
Begin VB.Form frmDeathMatch 
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   3915
   ClientTop       =   3630
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmDeathMatch.frx":0000
   ScaleHeight     =   3150
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmb_Clases 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   285
      ItemData        =   "frmDeathMatch.frx":258A
      Left            =   360
      List            =   "frmDeathMatch.frx":258C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lbl_CostoEntrada 
      BackStyle       =   0  'Transparent
      Caption         =   "30000"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1290
      Width           =   495
   End
   Begin VB.Label lbl_CostoMuerte 
      BackStyle       =   0  'Transparent
      Caption         =   "1500"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1545
      Width           =   615
   End
   Begin VB.Label lbl_PagaKill 
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lbl_Cerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblEntrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblAbandonar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   480
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2205
      Width           =   1815
   End
End
Attribute VB_Name = "frmDeathMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.MouseIcon = frmMain.MouseIcon
With cmb_Clases
    .AddItem "Mago"
    .AddItem "Clerigo"
    .AddItem "Paladin"
    .AddItem "Guerrero"
    .AddItem "Nigromante"
    .AddItem "Bardo"
    .AddItem "Druida"
    .AddItem "Asesino"
    .AddItem "Cazador"
    .AddItem "Arquero"
End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then
    Dim res As Long
    ReleaseCapture
    res = SendMessage(Me.Hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End If
End Sub

Private Sub lbl_Cerrar_Click()
Me.Visible = False
End Sub

Private Sub lblAbandonar_Click()
If MsgBox("¿Estás seguro de abandonar?", vbYesNo, "¡Atención!") = vbYes Then
Call SendData("/ABANDONARDM")
Me.Visible = False
End If

End Sub

Private Sub lblEntrar_Click()
If Val(UserGLD) < Val(lbl_CostoEntrada) Then
    AddtoRichTextBox frmMain.RecTxt, "DeathMatch> No tenés suficiente oro.", 2, 51, 223, 1, 1
    Me.Visible = False
    Exit Sub
End If

Call SendData("XDM " & UCase$(cmb_Clases.List(cmb_Clases.ListIndex)))
Me.Visible = False
End Sub
