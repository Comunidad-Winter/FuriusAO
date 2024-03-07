VERSION 5.00
Begin VB.Form FrmBPanel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   1650
   ClientTop       =   3285
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   Picture         =   "FrmBPanel.frx":0000
   ScaleHeight     =   4560
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.Image Comandos 
      Height          =   375
      Index           =   4
      Left            =   4440
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Comandos 
      Height          =   255
      Index           =   3
      Left            =   1560
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Image Comandos 
      Height          =   255
      Index           =   2
      Left            =   1680
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Comandos 
      Height          =   255
      Index           =   1
      Left            =   1800
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label OroLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Comandos 
      Height          =   255
      Index           =   0
      Left            =   1680
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "FrmBPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Comandos_Click(Index As Integer)
'0 = Transferir ORO
'1 = Retirar oro
'2 = Depositar oro
'3 = Ver boveda
'4 = Cierra boveda


Select Case Index
Case 0
frmTransferencia.Show , frmMain
Call frmTransferencia.GoMe
Me.Hide
Call SendData("#;")

Case 1
Dim ValA As Double
ValA = Val(InputBox("Ingrese la cantidad de dinero a retirar de la cuenta.", "Banco"))
If ValA > 0 Then
Call SendData("#0 " & ValA)
End If

Case 2
Dim ValB As Double
ValB = Val(InputBox("Ingrese la cantidad de dinero a depositar.", "Banco"))
If ValB > 0 Then
Call SendData("#Ñ " & ValB)
End If

Case 3
Call SendData("#W")

Case 4
Me.Hide
Call SendData("#;")

End Select
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

