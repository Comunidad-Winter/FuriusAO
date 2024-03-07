VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmIntro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer BuscarEngine 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin InetCtlsObjects.Inet Ipxd 
      Left            =   600
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Actualizo 
      Left            =   0
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":0614
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":091E
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   1200
      MouseIcon       =   "FrmIntro.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   3135
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BuscarEngine_Timer()
If frmPres.Visible = True Then frmPres.Hide
EnumTopWindows
End Sub

Private Sub Form_Load()
BuscarEngine.Enabled = True
BuscarEngine.Interval = 3000

Me.Picture = LoadPicture(App.Path & "\Graficos\MenuRapido.gif")

Dim corriendo As Integer
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim pepe As String

Dim exename As String
snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)
i = 0
While theloop <> 0
    exename = proc.szExeFile
    Text1.Text = proc.szExeFile
    If Text1.Text = "FuriusAO.exe" Or Text1.Text = "FúriusAO.exe" Then
        corriendo = corriendo + 1
        Text1.Text = ""
    End If
    i = i + 1
    theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap

End Sub



Private Sub Image2_Click()
Call Main
End Sub

Private Sub Image3_Click()
Shell App.Path & "/AOSETUP.EXE"
End Sub

Private Sub Image4_Click()
ShellExecute Me.Hwnd, "open", "http://www.furiusao.com.ar/Manual", "", "", 1


End Sub

Private Sub Image5_Click()
ShellExecute Me.Hwnd, "open", "http://www.furiusao.com.ar", "", "", 1

End Sub

Private Sub Image6_Click()
Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton Then

      dX = X

      dy = Y

      bmoving = True

   End If

   

End Sub

 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving And ((X <> dX) Or (Y <> dy)) Then

      Move Left + (X - dX), Top + (Y - dy)

   End If

   

End Sub

 

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then

      bmoving = False

   End If

   

End Sub

