VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Spawn"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
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
      Height          =   2370
      ItemData        =   "frmSpawnList.frx":0000
      Left            =   480
      List            =   "frmSpawnList.frx":0002
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Image command2 
      Height          =   255
      Left            =   2760
      MouseIcon       =   "frmSpawnList.frx":0004
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   615
   End
   Begin VB.Image command1 
      Height          =   495
      Left            =   960
      MouseIcon       =   "frmSpawnList.frx":030E
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call SendData("SPA" & lstCriaturas.ListIndex + 1)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    DX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> DX) Or (Y <> dy)) Then Move Left + (X - DX), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "SpawnList.gif")

End Sub

