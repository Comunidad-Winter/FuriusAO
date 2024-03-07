VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas de paz"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4725
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   580
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmPeaceProp.frx":0004
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmPeaceProp.frx":030E
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3120
      MouseIcon       =   "frmPeaceProp.frx":0618
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ParsePeaceOffers(ByVal s As String)
Dim T%, R%

T% = Val(ReadField(1, s, 44))

For R% = 1 To T%
    Call Lista.AddItem(ReadField(R% + 1, s, 44))
Next R%

Me.Show vbModeless, frmMain

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "OfertaDePazParaGuildMaster.gif")

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    DX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> DX) Or (Y <> dy)) Then Move Left + (X - DX), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
Private Sub Image1_Click()
Call SendData("ACEPPEAT" & Lista.List(Lista.ListIndex))
Unload Me
End Sub

Private Sub Image2_Click()
Call SendData("PEACEDET" & Lista.List(Lista.ListIndex))
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

