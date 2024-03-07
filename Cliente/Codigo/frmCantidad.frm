VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
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
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   375
      Left            =   570
      MaxLength       =   7
      TabIndex        =   0
      Top             =   790
      Width           =   2535
   End
   Begin VB.Image Command2 
      Height          =   330
      Left            =   1790
      MouseIcon       =   "frmCantidad.frx":0000
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   1280
      Width           =   1270
   End
   Begin VB.Image Command1 
      Height          =   330
      Left            =   600
      MouseIcon       =   "frmCantidad.frx":030A
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   1280
      Width           =   1140
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Sub Command1_Click()
frmCantidad.Visible = False
Call SendData("TI" & ItemElegido & "," & frmCantidad.Text1.Text)
frmCantidad.Text1.Text = "0"

End Sub
Private Sub Command2_Click()

frmCantidad.Visible = False

If ItemElegido <> FLAGORO Then
    Call SendData("TI" & ItemElegido & "," & UserInventory(ItemElegido).Amount)
Else: Call SendData("TI" & ItemElegido & "," & UserGLD)
End If

frmCantidad.Text1.Text = "0"

End Sub



Private Sub Form_Deactivate()

Unload Me

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "WinTirar.gif")

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    dX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> dX) Or (Y <> dy)) Then Call Move(Left + (X - dX), Top + (Y - dy))

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Text1_Change()

If Val(Text1.Text) < 0 Then
    Text1.Text = MAX_INVENTORY_OBJS
End If

If Val(Text1.Text) > MAX_INVENTORY_OBJS And ItemElegido <> FLAGORO Then
    Text1.Text = 1
End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 8) Then
    If (index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End If

End Sub

