VERSION 5.00
Begin VB.Form frmTorneo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Torneos"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
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
   ScaleHeight     =   4725
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListIns 
      Height          =   645
      Left            =   3240
      TabIndex        =   19
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtguildnews 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   3360
      Width           =   5775
   End
   Begin VB.CommandButton COT 
      Caption         =   "Organizar Torneo"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Inscribirse"
      Height          =   375
      Left            =   2040
      MouseIcon       =   "frmTorneo.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox TXTPrecio 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   15
      Text            =   "0"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox TXTPjs 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "2"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox TXTPR 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "50"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   "Modo"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2895
      Begin VB.OptionButton Modo2 
         Caption         =   "TODOS Vs TODOS"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Modo1 
         Caption         =   "1 Vs 1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Valido"
      Height          =   1455
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   2655
      Begin VB.OptionButton Val2 
         Caption         =   "Vale Todo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Val1 
         Caption         =   "Sin Invi, y sin todas las que modifiquen la jugabilidad."
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Paralizar siempre esta habilitado"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmTorneo.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label TXTLider 
      Caption         =   "Nadie"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Lider:"
      Height          =   195
      Left            =   3240
      TabIndex        =   21
      Top             =   3000
      Width           =   405
   End
   Begin VB.Label Label5 
      Caption         =   "Inscriptos"
      Height          =   255
      Left            =   3240
      TabIndex        =   20
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion del Torneo."
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "% / recaudado"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Precio Inscripcion"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "MAX Pjs"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Premio Ganador"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label TXTNombreT 
      Caption         =   "No hay torneos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Call SendData("INSC")
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub COT_Click()
SendData "ATOR"
Do While PuedeTorneo = 0 And noaprobado = 0
    DoEvents
Loop
If PuedeTorneo = 1 Then
    Unload Me
    frmTorneoCrear.Show
End If
End Sub

