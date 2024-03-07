VERSION 5.00
Begin VB.Form frmfidelidad 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2040
      MouseIcon       =   "frmfidelidad.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   840
      MouseIcon       =   "frmfidelidad.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmfidelidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

If Fide = 1 Then
Me.Picture = LoadPicture(DirGraficos & "fidelidadrey.gif")
ElseIf Fide = 2 Then
Me.Picture = LoadPicture(DirGraficos & "fidelidadthek.gif")
Else
Unload Me
End If

End Sub

Private Sub Image1_Click()
Call SendData("SF" & Fide)
Unload FrmElegirCamino
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
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
