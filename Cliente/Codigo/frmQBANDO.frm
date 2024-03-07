VERSION 5.00
Begin VB.Form frmQUESTB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de usuarios a participar."
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Summonear"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Summonear"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Summonear"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
   End
   Begin VB.ListBox List3 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   3570
      Left            =   5400
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3570
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8040
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label cri 
      Caption         =   "0"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   240
      Width           =   375
   End
   Begin VB.Label ciu 
      Caption         =   "0"
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
   Begin VB.Label neu 
      Caption         =   "0"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Criminales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Ciudadanos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Neutrales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmQUESTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("/SUM " & List1.List(List1.ListIndex))
End Sub

Private Sub Command2_Click()
Call SendData("/SUM " & List2.List(List2.ListIndex))
End Sub

Private Sub Command3_Click()
Call SendData("/SUM " & List3.List(List3.ListIndex))
End Sub

Private Sub Form_Load()
neu.Caption = List1.ListCount
ciu.Caption = List2.ListCount
cri.Caption = List3.ListCount
End Sub

