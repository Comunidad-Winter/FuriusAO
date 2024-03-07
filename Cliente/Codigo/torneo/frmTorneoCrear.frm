VERSION 5.00
Begin VB.Form frmTorneoCrear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de un Torneo"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTNombreTorneo 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      MaxLength       =   30
      TabIndex        =   19
      Text            =   "Torneo"
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Info de Torneo"
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox CheckIT 
         Caption         =   "Inscribirse en el Torneo"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox TXTPrecio 
         Alignment       =   2  'Center
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
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "500"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TXTPjs 
         Alignment       =   2  'Center
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
         TabIndex        =   11
         Text            =   "14"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Modo"
         Height          =   855
         Left            =   3120
         TabIndex        =   8
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton Modo1 
            Caption         =   "1 Vs 1"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Modo2 
            Caption         =   "TODOS Vs TODOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Valido"
         Height          =   1455
         Left            =   3120
         TabIndex        =   5
         Top             =   960
         Width           =   2655
         Begin VB.OptionButton Val1 
            Caption         =   "Sin Invi, y todas las que modifiquen la jugabilidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton Val2 
            Caption         =   "Vale Todo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Paralizar siempre esta habilitado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.TextBox TXTPR 
         Alignment       =   2  'Center
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
         TabIndex        =   4
         Text            =   "50"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TXTGR 
         Alignment       =   2  'Center
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
         TabIndex        =   3
         Text            =   "50"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Precio Inscripcion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "MAX Pjs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Premio Ganador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "% / recaudado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Ganancias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "% / recaudado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmTorneoCrear.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   5040
      MouseIcon       =   "frmTorneoCrear.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del Torneo:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "frmTorneoCrear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PIN As Byte
Private Sub CheckIT_Click()
If PIN = 1 Then
PIN = 0
Else
PIN = 1
End If
End Sub

Private Sub TXTPjs_Change()
TXTPjs = Val(TXTPjs)
If TXTPjs < 2 Then TXTPjs = 2
If TXTPjs > 26 Then TXTPjs = 26
End Sub
Private Sub TXTPR_Change()
TXTPR = Val(TXTPR)
If Val(TXTPR) > 100 Then
    TXTPR = 100
End If
End Sub
Private Sub TXTGR_Change()
TXTGR = Val(TXTGR)
If Val(TXTGR) > 100 Then
    TXTGR = 100
End If
End Sub
Private Sub TXTPrecio_Change()
On Error Resume Next
If TXTPrecio = "GRATIS" Then
    Exit Sub
End If
If TXTPrecio <> Val(TXTPrecio) Then TXTPrecio = 1
TXTPrecio = Val(TXTPrecio)
If TXTPrecio = "" Then TXTPrecio = 1
If TXTPrecio > 5000 Then TXTPrecio = 5000
If TXTPrecio <= 0 Then TXTPrecio = "GRATIS"
End Sub



Private Sub Command1_Click()
If Len(TXTNombreTorneo.Text) <= 30 Then
    If Not AsciiValidos(TXTNombreTorneo) Then
        MsgBox "Nombre invalido."
        Exit Sub
    End If
Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub
End If
If Val(TXTPR) + Val(TXTGR) <> 100 Then
    MsgBox "La division de ganancias es invalida."
    Exit Sub
End If
Call SendData("CTOR" & TXTNombreTorneo & "," & Val(TXTPrecio) & "," & TXTPjs & "," & TXTPR & "," & TXTGR & "," & IIf(Val1.value = True, "0", "1") & "," & IIf(Modo1.value = True, "0", "1") & "," & PIN)

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

