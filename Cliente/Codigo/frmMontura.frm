VERSION 5.00
Begin VB.Form frmMontura 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                               Montura"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar nombre"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Experi 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   "Experiencia"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label MaxHitMag 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label MinHitMag 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label MinDefMag 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label MaxDefMag 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label MaxDef 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label MinDef 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "Daño magica maxima"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      Caption         =   "Daño magica minima"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   "Defensa magica maxima"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Defensa magica minima"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Defensa maxima"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Defensa minima"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Estilomascota 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Estilo"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Label MascLvl 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label LblLevel 
      BackColor       =   &H80000007&
      Caption         =   "Nivel"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Nmascota 
      BackColor       =   &H80000007&
      Caption         =   "nummasco"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label NombreMontura 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Label1"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label MaxHit 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label MinHit 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblMaxHit 
      BackColor       =   &H80000007&
      Caption         =   "Mayor Golpe"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblMinHit 
      BackColor       =   &H80000007&
      Caption         =   "Menor Golpe"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "frmMontura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then Exit Sub
If Text1.Text = NombreMontura.Caption Then Exit Sub
If Len(Text1.Text) > 15 Then
MsgBox ("No puedes superar los 15 caracteres")
Exit Sub
End If
Call SendData("NewNam" & Text1.Text & "," & Nmascota)
End Sub

Private Sub Form_Load()
NombreMontura.Caption = MonturaName
End Sub

Private Sub NombreMontura_Change()
Text1.Text = NombreMontura.Caption
End Sub

