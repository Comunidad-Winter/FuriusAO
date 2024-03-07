VERSION 5.00
Begin VB.Form frmGuildSol 
   BorderStyle     =   0  'None
   Caption         =   "Ingreso"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4725
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000004&
      Height          =   1215
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Image command2 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmSolicitud.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   615
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   1680
      MouseIcon       =   "frmSolicitud.frx":030A
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSolicitud.frx":0614
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   540
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CName As String

Private Sub command1_Click()

Dim f$

f$ = "SOLICITUD" & CName
f$ = f$ & "," & Replace(Text1, vbCrLf, "º")

Call SendData(f$)

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub RecieveSolicitud(ByVal GuildName As String)

CName = GuildName

End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "IngresoClan.gif")
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
