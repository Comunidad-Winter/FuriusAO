VERSION 5.00
Begin VB.Form frmTorneosLider 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de torneos"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
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
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5160
      MouseIcon       =   "frmTorneos.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Estadisticas de US"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "PreInscripciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      Begin VB.ListBox PreMembers 
         Height          =   1815
         ItemData        =   "frmTorneos.frx":0152
         Left            =   120
         List            =   "frmTorneos.frx":0154
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar al Concursante"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Comenzar Torneo"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmTorneos.frx":0156
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame txtnews 
      Caption         =   "Descripcion del Torneo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   5895
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmTorneos.frx":02A8
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Concursantes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ListBox Members 
         Height          =   1815
         ItemData        =   "frmTorneos.frx":03FA
         Left            =   120
         List            =   "frmTorneos.frx":03FC
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmTorneosLider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If PreMembers.ListIndex >= 0 Then
    Call SendData("INST" & PreMembers.ListIndex + 1)
    Command1.Enabled = False
    PreMembers.Enabled = False
End If
End Sub
Public Sub Aprobado()
    Call Members.AddItem(PreMembers.List(PreMembers.ListIndex))
    Dim LoopC As Integer
    Dim LoopCont As Integer
    Dim Lista(100) As String
    For LoopC = 0 To PreMembers.ListCount - 1
        If PreMembers.ListIndex <> LoopC Then
            Lista(LoopC) = PreMembers.List(LoopC)
        Else
            Lista(LoopC) = "//*"
        End If
    Next LoopC
    LoopCont = PreMembers.ListCount - 1
    PreMembers.Clear
    If LoopCont > 0 Then
        For LoopC = 0 To LoopCont
            If Lista(LoopC) <> "//*" Then Call PreMembers.AddItem(Lista(LoopC))
        Next LoopC
    End If
    Command1.Enabled = True
    PreMembers.Enabled = True
End Sub
Private Sub Command2_Click()
On Error Resume Next
Call SendData("TRUN")
Unload Me
frmMain.SetFocus
End Sub

Private Sub Command3_Click()
SendData "TACT" & txtguildnews
End Sub

Private Sub Command4_Click()
If PreMembers.ListIndex >= 0 Then Call SendData("ESTT" & PreMembers.ListIndex + 1)
End Sub

Private Sub Command8_Click()
On Error Resume Next
If Command1.Enabled = False Then Exit Sub
Unload Me
frmMain.SetFocus
End Sub
