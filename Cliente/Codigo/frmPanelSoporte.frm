VERSION 5.00
Begin VB.Form frmPanelSoporte 
   Caption         =   "Soporte Actual:"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Buscar > En caso de bug"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3050
      Width           =   2055
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Actualizar Lista!"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdResp 
      BackColor       =   &H000000FF&
      Caption         =   "Responder"
      Height          =   255
      Left            =   2280
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar Soporte"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtRespuesta 
      Height          =   1095
      Left            =   2160
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtSoporte 
      Height          =   1575
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.ListBox lstSoportes 
      Height          =   2790
      ItemData        =   "frmPanelSoporte.frx":0000
      Left            =   0
      List            =   "frmPanelSoporte.frx":0007
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Respondido:"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape shpResp 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Soporte a responder:"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Soporte recibido:"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lbl1 
      Caption         =   "Lista de Usuarios:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.Menu MnuP 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSum 
         Caption         =   "Traer"
      End
      Begin VB.Menu mnuIr 
         Caption         =   "Ir"
      End
      Begin VB.Menu mnuCarcel 
         Caption         =   "Carcel 40 Hierro"
         Index           =   0
      End
      Begin VB.Menu mnuCarcel 
         Caption         =   "Carcel 30 Hierro"
         Index           =   1
      End
      Begin VB.Menu mnuCarcel 
         Caption         =   "Carcel 20 Hierro"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmPanelSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Me.Hide
End Sub

Private Sub cmdEliminar_Click()
'Dim UserNick As String
'UserNick = InputBox("Ingrese en este cuadro el nick a borrar.", "!!!")
If MsgBox("Estás seguro de borrar este S.O.S?  " & UCase$(ReadField$(2, Me.Caption, Asc(":"))), vbYesNo) = vbNo Then Exit Sub
Call SendData("/BORSO " & UCase$(ReadField$(2, Me.Caption, Asc(":"))))
End Sub

Private Sub cmdFind_Click()
MsgBox "Este botón se utiliza en caso de que alguien no pueda enviar SOS y el SOS no pueda ser respondido.", vbOKOnly
Dim UserNick As String
UserNick = InputBox("Ingrese en este cuadro el nick a buscar.", "!!!")
Call SendData("/SOSDE " & UCase$(UserNick))
Me.Caption = "Soporte Actual:" & UCase$(UserNick)
End Sub

Private Sub cmdResp_Click()
shpResp.BackColor = vbGreen
Call SendData("/RESOS " & Right$(frmPanelSoporte.Caption, Len(frmPanelSoporte.Caption) - 15) & ";" & txtRespuesta)
End Sub

Private Sub cmdUpdate_Click()
frmPanelSoporte.Hide
Call SendData("/DAMESOS")
End Sub

Private Sub lstSoportes_DblClick()
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/SOSDE " & lstSoportes.List(lstSoportes.ListIndex))
Me.Caption = "Soporte Actual:" & lstSoportes.List(lstSoportes.ListIndex)
Me.txtRespuesta = ""
End Sub

Private Sub lstSoportes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu MnuP
End If
End Sub

Private Sub mnuCarcel_Click(Index As Integer)
If lstSoportes.ListIndex = -1 Then Exit Sub
Select Case Index
Case 0
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/CARCEL SOPORTE INVALIDO" & lstSoportes.List(lstSoportes.ListIndex) & "@40")
Case 1
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/CARCEL SOPORTE INVALIDO" & lstSoportes.List(lstSoportes.ListIndex) & "@30")
Case 2
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/CARCEL SOPORTE INVALIDO@" & lstSoportes.List(lstSoportes.ListIndex) & "@20")
End Select
End Sub

Private Sub mnuIr_Click()
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/IRA " & lstSoportes.List(lstSoportes.ListIndex))
End Sub

Private Sub mnuSum_Click()
If lstSoportes.ListIndex = -1 Then Exit Sub
Call SendData("/SUM " & lstSoportes.List(lstSoportes.ListIndex))
End Sub

Private Sub txtSoporte_Click()
txtRespuesta.SetFocus
End Sub

