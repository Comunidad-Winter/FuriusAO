VERSION 5.00
Begin VB.Form frmSOPORTE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soporte FúriusAo. "
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   13440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   13440
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton EnviaResp 
      Caption         =   "Responder"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox respuesta 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   0
      MaxLength       =   90
      TabIndex        =   1
      Top             =   2280
      Width           =   8295
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0E0FF&
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13380
   End
   Begin VB.Label Label1 
      Caption         =   "Escribe aqui la respuesta rapida, no debe superar los 100 caracteres."
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmSOPORTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & ";" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1
End If
End Sub

Private Sub Command1_Click()
Me.Visible = False
List1.Clear
End Sub

Private Sub EnviaResp_Click()
SendData ("RR" & ReadField(1, List1.List(List1.ListIndex), Asc(";")))
SendData ("CJ" & respuesta.Text)
End Sub

Private Sub Form_Deactivate()
Me.Visible = False
List1.Clear
End Sub

Private Sub Form_Load()
'frmMSG.Picture = LoadPicture(DirGraficos & "menu-fondo.jpg")
'List1.Clear





End Sub

Private Sub HScroll1_Change()
List1.Left = -HScroll1.value

End Sub

Private Sub List1_Click()
Dim ind As Integer
ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc(";")))
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu menU_usuario
End If

End Sub

Private Sub mnuBorrar_Click()
If List1.ListIndex < 0 Then Exit Sub
SendData ("SOSDONE" & List1.List(List1.ListIndex))

List1.RemoveItem List1.ListIndex

End Sub

Private Sub mnuIR_Click()
SendData ("/IRA " & ReadField(1, List1.List(List1.ListIndex), Asc(";")))
End Sub

Private Sub mnutraer_Click()
SendData ("/SUM " & ReadField(1, List1.List(List1.ListIndex), Asc(";")))
End Sub

Private Sub VScroll1_Change()
List1.Top = -VScroll1.value
End Sub

