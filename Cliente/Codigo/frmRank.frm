VERSION 5.00
Begin VB.Form frmRank 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "  "
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "frmRank"
   NegotiateMenus  =   0   'False
   Picture         =   "frmRank.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2175
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmRank.frx":7E91
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
End
Attribute VB_Name = "frmRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
Text1.Text = FrmIntro.Ipxd.OpenURL("http://ao.localstrike.com.ar/NOTICIAS.TXT")
End Sub


Private Sub Label1_Click()
Me.Visible = False
End Sub


