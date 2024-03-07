VERSION 5.00
Begin VB.Form frmSastre 
   BorderStyle     =   0  'None
   Caption         =   "Sastre"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   -90
   ClientWidth     =   5250
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
   ScaleHeight     =   4350
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "1"
      Top             =   3560
      Width           =   1695
   End
   Begin VB.ListBox lstRopas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   4065
   End
   Begin VB.Image command4 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmSastre.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   855
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmSastre.frx":030A
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmSastre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command3_Click()
On Error Resume Next
Dim stxtCantBuffer As String
stxtCantBuffer = txtCantidad.Text

Call SendData("SCR" & ObjSastre(lstRopas.ListIndex) & " " & stxtCantBuffer)

Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()

Me.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "Sastreria.gif")

End Sub

Private Sub txtCantidad_Change()
If Val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

If Val(txtCantidad.Text) > MAX_INVENTORY_OBJS Then
    txtCantidad.Text = 1
End If
If Not IsNumeric(txtCantidad.Text) Then txtCantidad.Text = "1"

End Sub
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
