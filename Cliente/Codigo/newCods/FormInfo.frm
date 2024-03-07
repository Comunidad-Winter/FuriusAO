VERSION 5.00
Begin VB.Form FormInfo 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   ClientHeight    =   2595
   ClientLeft      =   7935
   ClientTop       =   510
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   Picture         =   "FormInfo.frx":0000
   ScaleHeight     =   2595
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.Label ExpTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2460
      TabIndex        =   7
      Top             =   815
      Width           =   255
   End
   Begin VB.Label cantidadmana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2460
      TabIndex        =   6
      Top             =   1625
      Width           =   255
   End
   Begin VB.Label cantidadhambre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2460
      TabIndex        =   5
      Top             =   1890
      Width           =   255
   End
   Begin VB.Label cantidadsta 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2460
      TabIndex        =   4
      Top             =   1090
      Width           =   255
   End
   Begin VB.Label cantidadagua 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2460
      TabIndex        =   3
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label cantidadhp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2460
      TabIndex        =   2
      Top             =   1370
      Width           =   255
   End
   Begin VB.Label LblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label LvlLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   550
      Width           =   105
   End
   Begin VB.Shape ShapeExp 
      BackColor       =   &H009DCAE7&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      Height          =   180
      Left            =   1635
      Top             =   825
      Width           =   1995
   End
   Begin VB.Shape Hpshp 
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   1635
      Top             =   1360
      Width           =   1995
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   175
      Left            =   1635
      Top             =   1100
      Width           =   1995
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   175
      Left            =   1635
      Top             =   1645
      Width           =   1995
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   175
      Left            =   1635
      Top             =   1910
      Width           =   1995
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   175
      Left            =   1630
      Top             =   2175
      Width           =   2000
   End
End
Attribute VB_Name = "FormInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub moverForm()
    Dim res As Long
    '
    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub

Private Sub Form_Click()
frmMain.SetFocus
End Sub

Private Sub Form_DblClick()
frmMain.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then moverForm
End Sub
