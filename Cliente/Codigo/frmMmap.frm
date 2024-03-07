VERSION 5.00
Begin VB.Form frmMmap 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicMMAp 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.Shape ShpPos 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         Height          =   90
         Left            =   1208
         Top             =   1208
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmMmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub moverForm()
    Dim res As Long
    '
    ReleaseCapture
   res = SendMessage(Me.Hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub

Private Sub PicMMAp_DblClick()
Me.Hide
End Sub

Private Sub PicMMAp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then moverForm
End Sub

Sub MoverPJ()
ShpPos.Top = PicMMAp.Height / 100 * UserPos.Y
ShpPos.Left = PicMMAp.Width / 100 * UserPos.X


End Sub
