VERSION 5.00
Begin VB.Form frmMensaje 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Mensaje BroadCast"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   405
      Left            =   0
      MouseIcon       =   "frmMensaje.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label msg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   570
      TabIndex        =   0
      Top             =   840
      Width           =   2895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    dX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> dX) Or (Y <> dy)) Then Move Left + (X - dX), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Form_Deactivate()
Me.SetFocus
End Sub


Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "Broadcast.gif")

End Sub

Private Sub Image1_Click()
Unload Me
End Sub

