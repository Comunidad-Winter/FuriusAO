VERSION 5.00
Begin VB.Form frmSPT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FuriusAO 1.0"
   ClientHeight    =   3045
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   3045
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Height          =   1230
      ItemData        =   "frmSPT.frx":0000
      Left            =   0
      List            =   "frmSPT.frx":0002
      TabIndex        =   4
      Top             =   240
      Width           =   11895
   End
   Begin VB.CommandButton EnviaResp 
      Caption         =   "Responder"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
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
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   10815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Escribe una respuesta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   855
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
Attribute VB_Name = "frmSPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1
End If
End Sub

Private Sub borrarsos_Click()
If List1.ListIndex < 0 Then Exit Sub
SendData ("SOSDONE" & List1.List(List1.ListIndex))
SendData ("SOSCONE" & List2.List(List2.ListIndex))

List1.RemoveItem List1.ListIndex
List2.RemoveItem List2.ListIndex
End Sub

Private Sub Command1_Click()
Me.Visible = False
List1.Clear
List2.Clear
End Sub

Private Sub Command2_Click()

End Sub

Private Sub EnviaResp_Click()
 SendData ("RR" & List1.List(List1.ListIndex))
 SendData ("CJ" & respuesta.Text)
End Sub

Private Sub Form_Deactivate()
Me.Visible = False
List1.Clear
List2.Clear
End Sub

Private Sub Form_Load()
List1.Clear
List2.Clear
End Sub

Private Sub ir_Click()
Call SendData("/IRA " & List1.List(List1.ListIndex))
End Sub

Private Sub irinvi_Click()
Call SendData("/INVISIBLE")
Call SendData("/IRA " & List1.List(List1.ListIndex))
End Sub

Private Sub List1_Click()
Dim ind As Integer
ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc("-")))
List2.ListIndex = ind
End Sub





Private Sub traer_Click()
Call SendData("/SUM " & List1.List(List1.ListIndex))
End Sub

