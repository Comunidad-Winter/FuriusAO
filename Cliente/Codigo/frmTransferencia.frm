VERSION 5.00
Begin VB.Form frmTransferencia 
   BorderStyle     =   0  'None
   Caption         =   "Panel de Transferencias"
   ClientHeight    =   1635
   ClientLeft      =   1650
   ClientTop       =   3285
   ClientWidth     =   3960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTransferencia.frx":0000
   ScaleHeight     =   1635
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtNombre2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtCantidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3600
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2880
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Paso As Integer
Dim Nombre1  As String
Dim Nombre2 As String
Dim cantidad As Long
Private Sub cmdGo_Click()
Call cmdGoX
End Sub
Private Sub cmdGoX()
Select Case Paso
Case 1
If txtNombre1.Text = "" Then Exit Sub
Nombre1 = txtNombre1.Text
txtNombre1.Visible = False
txtCantidad.Visible = True
txtCantidad.Text = ""
Paso = 2
txtCantidad.SetFocus
Me.Picture = LoadPicture(DirGraficos & "/IngresarOro.gif")
'lblStatus.Caption = "Ingrese la cantidad a transferir."
Case 2
If Val(txtCantidad.Text) = 0 Then Exit Sub
    If Val(txtCantidad.Text) > 5000000 Then
    MsgBox "No se pueden transferir mas de 5000000 monedas de oro a la ves."
    Exit Sub
End If
cantidad = Val(txtCantidad.Text)
txtCantidad.Visible = False
txtNombre2.Visible = True
txtNombre2.SetFocus
txtNombre2.Text = ""
Paso = 3
Me.Picture = LoadPicture(DirGraficos & "/Nombre2oro.gif")
'lblStatus.Caption = "Ingrese nuevamente el nombre a quién desea transferir el dinero."
Case 3
If txtNombre2.Text = "" Then Exit Sub
Nombre2 = txtNombre2.Text
If UCase$(Nombre1) <> UCase$(Nombre2) Then
MsgBox "Error. Los nombres deben ser iguales."
Me.Hide
Exit Sub
End If
If MsgBox("¿Estás seguro de enviar " & cantidad & " monedas de oro a " & UCase$(Nombre1) & " ?", vbYesNo) = vbYes Then
Call SendData("TRANSF" & UCase$(Nombre1) & "@" & cantidad)
Me.Visible = False
End If
End Select
End Sub

Public Sub GoMe()
Paso = 1
txtNombre1.Visible = True
txtNombre2.Visible = False
txtCantidad.Visible = False
txtNombre1.Text = ""
txtNombre2.Text = ""
txtCantidad.Text = ""
Nombre1 = ""
Nombre2 = ""
cantidad = 0
'lblStatus.Caption = "Ingrese el nombre a quién transferir el dinero."
Me.Picture = LoadPicture(DirGraficos & "/NombreOro.gif")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdGoX
End Sub

Private Sub Image1_Click()
Call cmdGoX
End Sub

Private Sub Image2_Click()
Me.Hide
End Sub
