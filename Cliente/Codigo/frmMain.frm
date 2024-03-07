VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
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
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   120
      Top             =   2760
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer tmrDopa 
      Interval        =   1000
      Left            =   1560
      Top             =   3240
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   -120
      ScaleHeight     =   15
      ScaleWidth      =   135
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer tmrAntiSH 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4080
      Top             =   7560
   End
   Begin VB.Timer ContResp 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1080
      Top             =   3240
   End
   Begin VB.Timer SndFPS 
      Interval        =   10000
      Left            =   600
      Top             =   3240
   End
   Begin Captura.wndCaptura Captura1 
      Left            =   120
      Top             =   2280
      _ExtentX        =   688
      _ExtentY        =   688
   End
   Begin VB.Frame frInvent 
      BorderStyle     =   0  'None
      Height          =   4358
      Left            =   8520
      TabIndex        =   16
      Top             =   1680
      Width           =   3216
      Begin VB.Image Image5 
         Height          =   195
         Index           =   3
         Left            =   1515
         MouseIcon       =   "frmMain.frx":1982
         MousePointer    =   99  'Custom
         Top             =   3920
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   195
         Index           =   2
         Left            =   1515
         MouseIcon       =   "frmMain.frx":1C8C
         MousePointer    =   99  'Custom
         Top             =   3600
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   1
         Left            =   1730
         MouseIcon       =   "frmMain.frx":1F96
         MousePointer    =   99  'Custom
         Top             =   3700
         Width           =   200
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   0
         Left            =   1380
         MouseIcon       =   "frmMain.frx":22A0
         MousePointer    =   99  'Custom
         Top             =   3705
         Width           =   195
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   3240
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   1780
         TabIndex        =   51
         Top             =   1160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   1440
         TabIndex        =   39
         Top             =   850
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Height          =   480
         Index           =   3
         Left            =   1440
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   25
         Left            =   2740
         TabIndex        =   73
         Top             =   3320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   24
         Left            =   2260
         TabIndex        =   72
         Top             =   3320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   23
         Left            =   1780
         TabIndex        =   71
         Top             =   3320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   22
         Left            =   1300
         TabIndex        =   70
         Top             =   3320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   21
         Left            =   820
         TabIndex        =   69
         Top             =   3320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   16
         Left            =   820
         TabIndex        =   68
         Top             =   2780
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   17
         Left            =   1300
         TabIndex        =   67
         Top             =   2780
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   18
         Left            =   1780
         TabIndex        =   66
         Top             =   2780
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   19
         Left            =   2260
         TabIndex        =   65
         Top             =   2780
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   20
         Left            =   2740
         TabIndex        =   64
         Top             =   2780
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   15
         Left            =   2740
         TabIndex        =   63
         Top             =   2240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   14
         Left            =   2260
         TabIndex        =   62
         Top             =   2240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   13
         Left            =   1780
         TabIndex        =   61
         Top             =   2240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   12
         Left            =   1300
         TabIndex        =   60
         Top             =   2240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   11
         Left            =   820
         TabIndex        =   59
         Top             =   2240
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   2740
         TabIndex        =   58
         Top             =   1700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   9
         Left            =   2260
         TabIndex        =   57
         Top             =   1700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   1780
         TabIndex        =   56
         Top             =   1700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   1300
         TabIndex        =   55
         Top             =   1700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   820
         TabIndex        =   54
         Top             =   1700
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   2740
         TabIndex        =   53
         Top             =   1160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   2260
         TabIndex        =   52
         Top             =   1160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   1300
         TabIndex        =   50
         Top             =   1160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   820
         TabIndex        =   49
         Top             =   1160
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   41
         Top             =   850
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   480
         Stretch         =   -1  'True
         Top             =   870
         Width           =   480
      End
      Begin VB.Label lblHechizos 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1800
         MouseIcon       =   "frmMain.frx":25AA
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   1440
         TabIndex        =   34
         Top             =   1390
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   960
         TabIndex        =   40
         Top             =   850
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   1920
         TabIndex        =   38
         Top             =   850
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   2400
         TabIndex        =   37
         Top             =   850
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   36
         Top             =   1390
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   960
         TabIndex        =   35
         Top             =   1390
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   11
         Left            =   480
         TabIndex        =   31
         Top             =   1940
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   12
         Left            =   960
         TabIndex        =   30
         Top             =   1940
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   13
         Left            =   1440
         TabIndex        =   29
         Top             =   1940
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   14
         Left            =   1920
         TabIndex        =   28
         Top             =   1940
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   15
         Left            =   2400
         TabIndex        =   27
         Top             =   1940
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   480
         TabIndex        =   26
         Top             =   2470
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   17
         Left            =   960
         TabIndex        =   25
         Top             =   2470
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   18
         Left            =   1440
         TabIndex        =   24
         Top             =   2470
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   19
         Left            =   1920
         TabIndex        =   23
         Top             =   2470
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   20
         Left            =   2400
         TabIndex        =   22
         Top             =   2470
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   21
         Left            =   480
         TabIndex        =   21
         Top             =   3020
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   22
         Left            =   960
         TabIndex        =   20
         Top             =   3020
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   23
         Left            =   1440
         TabIndex        =   19
         Top             =   3020
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   24
         Left            =   1920
         TabIndex        =   18
         Top             =   3020
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   25
         Left            =   2400
         TabIndex        =   17
         Top             =   3020
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   1920
         TabIndex        =   33
         Top             =   1390
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   10
         Left            =   2400
         TabIndex        =   32
         Top             =   1390
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   960
         Stretch         =   -1  'True
         Top             =   870
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   4
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   870
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   5
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   870
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   6
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1410
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   7
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1410
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   8
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1410
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   9
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1410
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   10
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1410
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   11
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   12
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   13
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   14
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   15
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1950
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   16
         Left            =   480
         Stretch         =   -1  'True
         Top             =   2490
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   17
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2490
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   18
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2490
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   19
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   2490
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   20
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2490
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   21
         Left            =   480
         Stretch         =   -1  'True
         Top             =   3030
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   22
         Left            =   960
         Stretch         =   -1  'True
         Top             =   3030
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   23
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   3030
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   24
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   3030
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   25
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   3030
         Width           =   480
      End
      Begin VB.Image imgFondoInvent 
         Height          =   4380
         Left            =   0
         Top             =   0
         Width           =   3225
      End
   End
   Begin VB.Timer tmrBmp 
      Left            =   1560
      Top             =   2280
   End
   Begin VB.Timer trabajo 
      Enabled         =   0   'False
      Left            =   600
      Top             =   2760
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   1080
      Top             =   2760
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   600
      Top             =   2280
   End
   Begin VB.Timer FPS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   2280
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2040
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.Timer Attack 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   2760
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1680
      Visible         =   0   'False
      Width           =   7027
   End
   Begin VB.Frame frHechizos 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   8520
      TabIndex        =   43
      Top             =   1680
      Width           =   3240
      Begin VB.ListBox lstHechizos 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2565
         Left            =   420
         TabIndex        =   44
         Top             =   1095
         Width           =   2595
      End
      Begin VB.Label lblLanzar 
         BackStyle       =   0  'Transparent
         Height          =   360
         Left            =   390
         MouseIcon       =   "frmMain.frx":28B4
         MousePointer    =   99  'Custom
         TabIndex        =   88
         Top             =   3840
         Width           =   1305
      End
      Begin VB.Label lblCh 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   120
         TabIndex        =   87
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label lblInvent 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   405
         MouseIcon       =   "frmMain.frx":2BBE
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   15
         Width           =   1350
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Height          =   360
         Left            =   1965
         MouseIcon       =   "frmMain.frx":2EC8
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   3840
         Width           =   1050
      End
      Begin VB.Label lblAbajo 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         MouseIcon       =   "frmMain.frx":31D2
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   780
         Width           =   300
      End
      Begin VB.Label lblArriba 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2415
         MouseIcon       =   "frmMain.frx":34DC
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   780
         Width           =   300
      End
      Begin VB.Image imgFondoHechizos 
         Height          =   4395
         Left            =   0
         Top             =   0
         Width           =   3240
      End
   End
   Begin MSWinsockLib.Winsock WSAntiSH 
      Left            =   4680
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "time.windows.com"
      RemotePort      =   37
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1425
      Left            =   120
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2514
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":37E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ImgMen 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   11040
      MouseIcon       =   "frmMain.frx":3864
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":452E
      Top             =   1275
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblMSG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "¡Tienes un nuevo mensaje!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   10080
      TabIndex        =   86
      Top             =   1500
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image imgSoporte 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   11040
      Picture         =   "frmMain.frx":459A
      Top             =   1275
      Width           =   300
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "17.000.000/17.000.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   9960
      TabIndex        =   85
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "52,32%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   9960
      TabIndex        =   84
      Top             =   960
      Width           =   1695
   End
   Begin VB.Shape shpExp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400040&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   9675
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   2265
   End
   Begin VB.Label Agilidad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   255
      Left            =   7440
      TabIndex        =   83
      Top             =   8640
      Width           =   615
   End
   Begin VB.Shape shpCl 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Left            =   7440
      Top             =   8640
      Width           =   15
   End
   Begin VB.Label Fuerza 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8400
      TabIndex        =   82
      Top             =   8640
      Width           =   675
   End
   Begin VB.Shape shpFz 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Left            =   8400
      Top             =   8640
      Width           =   15
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   4
      Left            =   10545
      MouseIcon       =   "frmMain.frx":4604
      MousePointer    =   99  'Custom
      Top             =   6450
      Width           =   1170
   End
   Begin VB.Image Party 
      Height          =   315
      Left            =   10545
      MouseIcon       =   "frmMain.frx":52CE
      MousePointer    =   99  'Custom
      Top             =   8040
      Width           =   1170
   End
   Begin VB.Shape MainViewShp 
      BackColor       =   &H000080FF&
      BorderColor     =   &H00000000&
      Height          =   6225
      Left            =   120
      Top             =   1995
      Width           =   8190
   End
   Begin VB.Label NumOnline 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   6765
      TabIndex        =   79
      Top             =   8700
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   10560
      MouseIcon       =   "frmMain.frx":5F98
      MousePointer    =   99  'Custom
      TabIndex        =   78
      Top             =   1290
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   10800
      MouseIcon       =   "frmMain.frx":6C62
      MousePointer    =   99  'Custom
      TabIndex        =   77
      Top             =   1290
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6720
      TabIndex        =   76
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   10320
      MouseIcon       =   "frmMain.frx":792C
      MousePointer    =   99  'Custom
      TabIndex        =   75
      Top             =   1290
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label modo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "1 Normal"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   225
      TabIndex        =   74
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   1080
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label arma 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   2880
      TabIndex        =   1
      Top             =   8610
      Width           =   270
   End
   Begin VB.Label armadura 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   525
      TabIndex        =   15
      Top             =   8610
      Width           =   270
   End
   Begin VB.Label escudo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   3840
      TabIndex        =   14
      Top             =   8610
      Width           =   270
   End
   Begin VB.Label casco 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1680
      TabIndex        =   13
      Top             =   8610
      Width           =   270
   End
   Begin VB.Label mapa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada al Castillo Azul [190-50-50 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      MouseIcon       =   "frmMain.frx":85F6
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   8580
      Width           =   2775
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   9960
      MouseIcon       =   "frmMain.frx":92C0
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label cantidadhp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8970
      TabIndex        =   10
      Top             =   7350
      Width           =   1200
   End
   Begin VB.Label cantidadagua 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8970
      TabIndex        =   9
      Top             =   8055
      Width           =   1200
   End
   Begin VB.Label cantidadsta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8970
      TabIndex        =   11
      Top             =   6645
      Width           =   1200
   End
   Begin VB.Label cantidadhambre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8970
      TabIndex        =   8
      Top             =   7695
      Width           =   1200
   End
   Begin VB.Label cantidadmana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8970
      TabIndex        =   7
      Top             =   7005
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   11160
      MouseIcon       =   "frmMain.frx":AC42
      MousePointer    =   99  'Custom
      Top             =   75
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00003E25&
      X1              =   16
      X2              =   551.467
      Y1              =   126.333
      Y2              =   126.333
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   11520
      MouseIcon       =   "frmMain.frx":B90C
      MousePointer    =   99  'Custom
      Top             =   75
      Width           =   255
   End
   Begin VB.Label fpstext 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   6750
      TabIndex        =   6
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "aRankyr"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   9840
      TabIndex        =   5
      Top             =   540
      Width           =   1665
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      Height          =   105
      Left            =   8865
      Top             =   6690
      Width           =   1425
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   105
      Left            =   8865
      Top             =   7050
      Width           =   1425
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   10680
      TabIndex        =   4
      Top             =   6120
      Width           =   90
   End
   Begin VB.Shape Hpshp 
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   8865
      Shape           =   4  'Rounded Rectangle
      Top             =   7395
      Width           =   1425
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   105
      Left            =   8865
      Top             =   7740
      Width           =   1425
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   105
      Left            =   8865
      Top             =   8100
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   0
      Left            =   10560
      MouseIcon       =   "frmMain.frx":C5D6
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   1
      Left            =   10545
      MouseIcon       =   "frmMain.frx":D2A0
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   2
      Left            =   10545
      MouseIcon       =   "frmMain.frx":DF6A
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   1245
   End
   Begin VB.Label lblLvl2 
      BackStyle       =   0  'Transparent
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   345
      Left            =   9285
      TabIndex        =   3
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   8550
      MouseIcon       =   "frmMain.frx":EC34
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1125
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AMS As AntiMacros
Private Type BLENDFUNCTION
BlendOp As Byte
BlendFlags As Byte
SourceConstantAlpha As Byte
AlphaFormat As Byte
End Type
Private Const AC_SRC_OVER = &H0
   
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal xOriginDest As Long, ByVal yOriginDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hdcSrc As Long, ByVal xOriginSrc As Long, ByVal yOriginSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
   
Dim Blend As BLENDFUNCTION
Dim blendlong As Long
Dim Contador As Integer

Public Pocho As Boolean
Public MiMapitA As Integer

Public ActualSecond As Long
Public LastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long

Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim POS(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte
Public boton As Integer

Dim endEvent As Long
Implements DirectXEvent
Dim m_Jpeg As cJpeg
Dim PuedeLanzarLbl As Byte
Private Sub BuscarEngine_Timer()
EnumTopWindows
End Sub


Private Sub cantidadagua_Change()
FormInfo.cantidadagua.Caption = cantidadagua.Caption
End Sub


Private Sub cantidadhambre_Change()
FormInfo.cantidadhambre.Caption = cantidadhambre.Caption
End Sub

Private Sub cantidadhp_Change()
FormInfo.cantidadhp.Caption = cantidadhp.Caption
End Sub

Private Sub cantidadmana_Change()
FormInfo.cantidadmana.Caption = cantidadmana.Caption
End Sub

Private Sub cantidadsta_Change()
FormInfo.cantidadsta.Caption = cantidadsta.Caption
End Sub

Private Sub exp_Change()
FormInfo.ExpTotal.Caption = exp.Caption
End Sub


Private Sub Form_Activate()

If frmParty.Visible Then frmParty.SetFocus
If frmParty2.Visible Then frmParty2.SetFocus

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
AMS.ClickKeyDown KeyCode
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AMS.ClickRatonDown
boton = Button

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not AMS.ClickRatonUP Then Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set AMS = Nothing
frmMmap.Hide
End Sub

Private Sub GldLbl_Change()
FormInv.GldLbl.Caption = GldLbl.Caption
End Sub




Private Sub Image5_Click(Index As Integer)

If (ItemElegido <= 0 Or ItemElegido > MAX_INVENTORY_SLOTS) Then Exit Sub
If ItemElegido = 1 And Index = 0 Then Exit Sub
If ItemElegido = MAX_INVENTORY_SLOTS And Index = 1 Then Exit Sub
If ItemElegido < 6 And Index = 2 Then Exit Sub
If ItemElegido > MAX_INVENTORY_SLOTS - 5 And Index = 3 Then Exit Sub

Call SendData("ZI" & ItemElegido & "," & Index)

Select Case Index
    Case 0
        Shape1.Top = imgObjeto(ItemElegido - 1).Top
        Shape1.Left = imgObjeto(ItemElegido - 1).Left
        ItemElegido = ItemElegido - 1
    Case 1
        Shape1.Top = imgObjeto(ItemElegido + 1).Top
        Shape1.Left = imgObjeto(ItemElegido + 1).Left
        ItemElegido = ItemElegido + 1
    Case 2
        Shape1.Top = imgObjeto(ItemElegido - 5).Top
        Shape1.Left = imgObjeto(ItemElegido - 5).Left
        ItemElegido = ItemElegido - 5
    Case 3
        Shape1.Top = imgObjeto(ItemElegido + 5).Top
        Shape1.Left = imgObjeto(ItemElegido + 5).Left
        ItemElegido = ItemElegido + 5
End Select

End Sub

Private Sub Image7_Click()

End Sub

Private Sub Image9_Click()

End Sub

Private Sub imgFondoInvent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AMS.ClickRatonDown
End Sub

Private Sub ImgMen_Click()
Call SendData("/MISOPORTE")
lblMSG.Visible = False
ImgMen.Visible = False
End Sub

Private Sub imgSoporte_Click()
Call SendData("/MISOPORTE")
lblMSG.Visible = False
ImgMen.Visible = False
End Sub

Private Sub Label2_Click(Index As Integer)

If ItemElegido <> Index And UserInventory(Index).Name <> "Nada" Then
    Shape1.Visible = True
    Shape1.Top = imgObjeto(Index).Top
    Shape1.Left = imgObjeto(Index).Left
    ItemElegido = Index
End If

End Sub

Private Sub Label3_Click()

Call SendData("#N")

End Sub

Private Sub Label5_Click()

Call SendData("#!")

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

Call SendData("#O")

End Sub

Private Sub Label8_Change()
FormInfo.LblName.Caption = Label8.Caption
End Sub

Private Sub lblarriba_Click()

If lstHechizos.ListIndex < 1 Then Exit Sub

If lstHechizos.ListIndex >= 1 Then Call SendData("DESPHE" & 1 & "," & lstHechizos.ListIndex + 1)
lstHechizos.ListIndex = lstHechizos.ListIndex - 1

End Sub
Private Sub lblabajo_Click()

If lstHechizos.ListIndex > 33 Then Exit Sub

If lstHechizos.ListIndex <= 33 Then Call SendData("DESPHE" & 2 & "," & lstHechizos.ListIndex + 1)
lstHechizos.ListIndex = lstHechizos.ListIndex + 1

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MouseX = X
MouseY = Y

If MouseX > 696 And MouseX < frmMain.LvlLbl.Left + frmMain.LvlLbl.Width And MouseY > frmMain.LvlLbl.Top And MouseY < frmMain.LvlLbl.Top + frmMain.LvlLbl.Height Then
exp.Visible = True
frmMain.LvlLbl.Visible = False
Else
exp.Visible = False
frmMain.LvlLbl.Visible = True
End If
End Sub
Private Sub FX_Timer()
Dim N As Byte

If FX = 0 And RandomNumber(1, 150) < 12 Then
    N = RandomNumber(1, 45)
    Select Case N
        Case Is <= 15
            Call PlayWaveDS("22.wav")
        Case Is <= 30
            Call PlayWaveDS("21.wav")
        Case Is <= 35
            Call PlayWaveDS("28.wav")
        Case Is <= 40
            Call PlayWaveDS("29.wav")
        Case Is <= 45
            Call PlayWaveDS("34.wav")
    End Select
End If

End Sub
Private Sub imgObjeto_Click(Index As Integer)

If ItemElegido <> Index And UserInventory(Index).Name <> "Nada" Then
    Shape1.Visible = True
    Shape1.Top = imgObjeto(Index).Top
    Shape1.Left = imgObjeto(Index).Left
    ItemElegido = Index
End If

End Sub
Private Sub imgObjeto_DblClick(Index As Integer)

If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

If ItemElegido = Index Then

If Not NoPuedeUsarC Then
NoPuedeUsarC = True
Call SendData("USE" & ItemElegido)
'Else
'Debug.Print "no"
End If

End If
    
    
End Sub

Private Sub lblCh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PuedeLanzarLbl = 0
End Sub

Private Sub lblHechizos_Click()

Call PlayWaveDS(SND_CLICK)
frHechizos.Visible = True
frInvent.Visible = False

End Sub
Private Sub lblInvent_Click()

Call PlayWaveDS(SND_CLICK)
frInvent.Visible = True
frHechizos.Visible = False

End Sub

Private Sub lblLanzar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AMS.ClickRatonDown
End Sub

'Private Sub lblLanzar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Not AMS.ClickKeyUP(KeyCode) Then Exit Sub
'End Sub

Private Sub lblObjCant_Click(Index As Integer)

If ItemElegido <> Index And UserInventory(Index).Name <> "Nada" Then
    Shape1.Visible = True
    Shape1.Top = imgObjeto(Index).Top
    Shape1.Left = imgObjeto(Index).Left
    ItemElegido = Index
End If

End Sub
Private Sub lblObjCant_DblClick(Index As Integer)

If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

If Not NoPuedeUsarC Then
NoPuedeUsarC = True
If ItemElegido = Index Then Call SendData("USE" & ItemElegido)
End If


End Sub
Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub
Private Sub CreateEvent()

endEvent = DirectX.CreateEvent(Me)

End Sub


Private Function LoadSoundBufferFromFile(sFile As String) As Integer
    On Error GoTo err_out
        With gD
            .lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPOSITIONNOTIFY
            .lReserved = 0
        End With
        Set gDSB = DirectSound.CreateSoundBufferFromFile(DirSound & sFile, gD, gW)
        With POS(0)
            .hEventNotify = endEvent
            .lOffset = -1
        End With
        DirectX.SetEvent endEvent

        
    Exit Function

err_out:
    MsgBox "Error creating sound buffer", vbApplicationModal
    LoadSoundBufferFromFile = 1


End Function


Public Sub Play(ByVal nombre As String, Optional ByVal LoopSound As Boolean = False)
    If FX = 1 Then Exit Sub
    Call LoadSoundBufferFromFile(nombre)

    If LoopSound Then
        gDSB.Play DSBPLAY_LOOPING
    Else
        gDSB.Play DSBPLAY_DEFAULT
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If endEvent Then DirectX.DestroyEvent endEvent

If prgRun Then
    prgRun = False
    Cancel = 1
End If

End Sub
Public Sub StopSound()
On Local Error Resume Next

If Not gDSB Is Nothing Then
    gDSB.Stop
    gDSB.SetCurrentPosition 0
End If

End Sub
Private Sub FPS_Timer()



If logged And Not frmMain.Visible Then
    Unload frmConnect
    frmMain.Show
End If
    
End Sub
Private Sub Image2_Click()

Me.WindowState = vbMinimized

End Sub
Private Sub Image4_Click()

ItemElegido = FLAGORO
If UserGLD > 0 Then frmCantidad.Show

End Sub

Private Sub LvlLbl_Change()
FormInfo.LvlLbl.Caption = LvlLbl.Caption
Me.lblLvl2.Caption = UserLvl
End Sub

Private Sub LvlLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LvlLbl.Visible = False
exp.Visible = True
End Sub


Private Sub mapa_Click()
        If frmMmap.Visible = True Then
            frmMmap.Hide
        Else
            frmMmap.Show , frmMain
            Call frmOpciones.Transparencia(frmMmap.Hwnd, 190)
        End If
End Sub

Private Sub mapa_DblClick()
frmMapa.Visible = Not frmMapa.Visible
End Sub

Private Sub Party_Click()

frmParty.ListaIntegrantes.Clear
LlegoParty = False
Call SendData("PARINF")
Do While Not LlegoParty
    DoEvents
Loop
frmParty.Visible = True
frmParty.SetFocus
LlegoParty = False
            
End Sub

Private Sub RecTxt_GotFocus()
'On Error Resume Next
SendTxt.Visible = False
frmMain.SetFocus

End Sub

Private Sub RecTxt_KeyPress(KeyAscii As Integer)
frmMain.SetFocus
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call ProcesaEntradaCmd(stxtbuffer)
    stxtbuffer = ""
    frmMain.SendTxt.Text = ""
    frmMain.SendTxt.Visible = False
    KeyCode = 0
End If

End Sub

Private Sub SndFPS_Timer()
If FramesPerSec And logged And Not Saliendo Then
Call SendData("NL" & EncriptarFPS(FramesPerSec))
Exit Sub
End If
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
    ActualSecond = Mid$(time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = LastSecond Then End
    LastSecond = ActualSecond
End Sub





Private Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
           If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show
           End If
        End If
    End If

 
End Sub

Private Sub AgarrarItem()
    SendData "AG"
 
End Sub

Private Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then
    SendData "USA" & ItemElegido
    End If
   
End Sub
Public Sub EquiparItem()

If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "EQUI" & ItemElegido
        
End Sub





Private Sub lblLanzar_Click()

If lstHechizos.List(lstHechizos.ListIndex) <> "Nada" And TiempoTranscurrido(LastHechizo) >= IntervaloSpell And TiempoTranscurrido(Hechi) >= IntervaloSpell / 4 Then
    Call SendData("LH" & lstHechizos.ListIndex + 1)
    Call SendData("UK" & Magia)
End If

End Sub
Private Sub lblInfo_Click()
    Call SendData("INFS" & lstHechizos.ListIndex + 1)
End Sub
Private Sub Form_Click()

If Cartel Then Cartel = False

If Comerciando = 0 Then
    Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)
    If Abs(UserPos.Y - tY) > 6 Then Exit Sub
    If Abs(UserPos.X - tX) > 8 Then Exit Sub
    If EligiendoWhispereo Then
        Call SendData("WH" & tX & "," & tY)
        EligiendoWhispereo = False
        Exit Sub
    End If
    
    If UsingSkill = 0 Then
        SendData "LC" & tX & "," & tY
    Else
        frmMain.MousePointer = vbDefault
        If UsingSkill = Magia Then
            If PuedeLanzarLbl < 4 Then
               If (TiempoTranscurrido(LastHechizo) < IntervaloSpell Or TiempoTranscurrido(Hechi) < IntervaloSpell / 4) Then
                    Exit Sub
               Else: Hechi = Timer
               End If
               PuedeLanzarLbl = PuedeLanzarLbl + 1
            Else
                Exit Sub
            End If
        ElseIf UsingSkill = Proyectiles Then
            If (TiempoTranscurrido(LastFlecha) < IntervaloFlecha Or TiempoTranscurrido(Flecho) < IntervaloFlecha / 4) Then
                Exit Sub
            Else: Flecho = Timer
            End If
        End If
            Call SendData("WLC" & tX & "," & tY & "," & UsingSkill)
'        Call SendData("WLC" & tX & "," & tY & ",200")
        UsingSkill = 0
    End If
End If

If boton = vbRightButton Then Call SendData("/TELEPLOC")
boton = 0

End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub


Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton And Desplazar = True Then

      dX = X

      dy = Y

      bmoving = True

   End If

   

End Sub
Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> dX) Or (Y <> dy)) And Desplazar = True Then
    Move Left + (X - dX), Top + (Y - dy)
    MainViewRect.Left = 7 + (frmMain.Left / 15) + 32 * RenderMod.iImageSize
    MainViewRect.Top = 152 + (frmMain.Top / 15) + 32 * RenderMod.iImageSize
    MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)
End If
   
End Sub
Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And Desplazar = True Then
    bmoving = False
    MainViewRect.Left = 7 + (frmMain.Left / 15) + 32 * RenderMod.iImageSize
    MainViewRect.Top = 152 + (frmMain.Top / 15) + 32 * RenderMod.iImageSize
    MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Not AMS.ClickKeyUP(KeyCode) Then Exit Sub


If Not SendTxt.Visible Then

    Select Case KeyCode
    

    Case vbKeyF10:
            SendData "INFOMASCOTA"
    'Case vbKeyF9:
' If Shift = 0 Then Exit Sub

               ' Call SendMascBox(UserIndex)
    Case CustomKeys.BindedKey(eKeyType.mKeyFullScreen)
    Dim Modificador As Integer
    frmMain.Cls
        
    If Pocho = True Then
        Pocho = False
        RenderMod.iImageSize = -4
        Modificador = 120 'frmMain.PanelDer.ScaleWidth / 2
        frmMain.MainViewShp.Left = MainViewWidth - MainViewHeight
        frmMain.MainViewShp.Top = MainViewWidth - MainViewHeight
        frmMain.MainViewShp.Height = frmMain.ScaleHeight 'MainViewRect.Right
        frmMain.MainViewShp.Width = frmMain.ScaleWidth 'MainViewRect.Bottom
        Me.Picture = Nothing
  GoTo feo
    
    ElseIf Pocho = False Then
        Pocho = True
        RenderMod.iImageSize = 0
        Modificador = 0
        frmMain.MainViewShp.Left = 4
        frmMain.MainViewShp.Top = 150
        frmMain.MainViewShp.Height = 411
        frmMain.MainViewShp.Width = 547
        Me.Picture = LoadPicture(DirGraficos & "Principal.gif")
    End If
    
    
feo:
    Image1(0).Visible = Pocho
    Image1(1).Visible = Pocho
    Image1(2).Visible = Pocho
    Party.Visible = Pocho
    
    Image2.Visible = Pocho
    Image3.Visible = Pocho
    fpstext.Visible = Pocho
    RecTxt.Visible = Pocho
    imgFondoInvent.Visible = Pocho

    FormConsola.Visible = Not Pocho
    If FormConsola.Visible = True Then FormConsola.Show , frmMain

    FormBarInv.Visible = Not Pocho
    If FormBarInv.Visible = True Then FormBarInv.Show , frmMain
    
    FormInfo.Visible = False
    FormInv.Visible = False
    FormTalk.Visible = False
    FormListOpciones.Visible = False
    
    
    
    
    
    Marriba.Visible = Not Pocho
    Mabajo.Visible = Not Pocho
    Misquierda.Visible = Not Pocho
    Mderecha.Visible = Not Pocho
    If Marriba.Visible Then Marriba.Show , frmMain
    Marriba.Enabled = False
    If Mderecha.Visible Then Mderecha.Show , frmMain
    Mderecha.Enabled = False
    If Misquierda.Visible Then Misquierda.Show , frmMain
    Misquierda.Enabled = False
    If Mabajo.Visible Then Mabajo.Show , frmMain
    Mabajo.Enabled = False
    'menues
    
    frInvent.Visible = Pocho
    frHechizos.Visible = Pocho
    imgFondoHechizos.Visible = Pocho

    armadura.Visible = Pocho
    arma.Visible = Pocho
    casco.Visible = Pocho
    escudo.Visible = Pocho
    NumOnline.Visible = Pocho
    GldLbl.Visible = Pocho
    mapa.Visible = Pocho
    AGUAsp.Visible = Pocho
    MANShp.Visible = Pocho
    Hpshp.Visible = Pocho
    STAShp.Visible = Pocho
    COMIDAsp.Visible = Pocho
    cantidadsta.Visible = Pocho
    cantidadmana.Visible = Pocho
    cantidadhp.Visible = Pocho
    cantidadhambre.Visible = Pocho
    cantidadagua.Visible = Pocho
    Label8.Visible = Pocho
    LvlLbl.Visible = Pocho
    'Label6.Visible = Pocho
    'Label3.Visible = Pocho
    'Label5.Visible = Pocho
    'Label7.Visible = Pocho
    exp.Visible = Pocho
    'Image8.Visible = Pocho
    'Image9.Visible = Pocho
    'Fuerza.Visible = Pocho
    'Agilidad.Visible = Pocho
    
    
    Dim Bucle As Integer
    For Bucle = 1 To 25
    imgObjeto(Bucle).Visible = Pocho
    DoEvents
    Next Bucle
   
    frmMain.SetFocus
    bInvMod = True

'Loop principal!
'[CODE]:MatuX' se llama a esta funcion otra vez Fedex
    MainViewRect.Left = MainViewLeft + 32 * RenderMod.iImageSize + Modificador
    MainViewRect.Top = MainViewTop + (32 * RenderMod.iImageSize) * 1.1875
    MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

    MainDestRect.Left = ((TilePixelWidth * TileBufferSize) - TilePixelWidth) + 32 * RenderMod.iImageSize
    MainDestRect.Top = ((TilePixelHeight * TileBufferSize) - TilePixelHeight) + 32 * RenderMod.iImageSize
    MainDestRect.Right = (MainDestRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainDestRect.Bottom = (MainDestRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)
'[END]'
    Exit Sub

        Case vbKeyF:
            If FX = 1 Then
                FX = 0
                frmOpciones.PictureFxs.Picture = LoadPicture(DirGraficos & "tick1.gif")
                Call AddtoRichTextBox(frmMain.RecTxt, "Los sonidos han sido activados.", 255, 255, 255, 1, 0)
            Else
                FX = 1
                frmOpciones.PictureFxs.Picture = LoadPicture(DirGraficos & "tick2.gif")
                Call AddtoRichTextBox(frmMain.RecTxt, "Los sonidos han sido desactivados.", 255, 255, 255, 1, 0)
            End If
            
            
'        Case vbKeyM:
         Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
            If Not IsPlayingCheck Then
                Musica = 0
                Play_Midi
                frmOpciones.PictureMusica.Picture = LoadPicture(DirGraficos & "tick1.gif")
            Else
                Musica = 1
                frmOpciones.PictureMusica.Picture = LoadPicture(DirGraficos & "tick2.gif")
                Stop_Midi
            End If
            
            Case vbKeyJ:
       Call SendData("^~") ' mandamos un paquete "raro" para que no haga interferencia en el servidor
       Exit Sub

            
        'Case vbKeyA:
        Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
            Call AgarrarItem
       ' Case vbKeyC:
        
        
        'Case vbKeyE:
        Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
            Call EquiparItem
        'Case vbKeyF12:
        Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
            Nombres = Not Nombres
        'Case vbKeyZ:
        Case CustomKeys.BindedKey(eKeyType.mKeyUnlock)
            Call SendData("(A")
            
        'Case vbKeyD
        Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
            Call SendData("UK" & Domar)
            
        'Case vbKeyR:
        Case CustomKeys.BindedKey(eKeyType.mKeySteal)
            Call SendData("UK" & Robar)
    
        'Case vbKeyO:
        Case CustomKeys.BindedKey(eKeyType.mKeyHide)
            Call SendData("UK" & Ocultarse)
            
        'Case vbKeyT:
        Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
            Call TirarItem
            
        Case vbKey1:
            frmMain.modo = "1 Normal"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
            
        Case vbKey2:
            Call AddtoRichTextBox(frmMain.RecTxt, "Has click sobre el usuario al que quieres susurrar.", 255, 255, 255, 1, 0)
            frmMain.modo = "2 Susurrar"
            MousePointer = 2
            EligiendoWhispereo = True
            
        Case vbKey3:
            frmMain.modo = "3 Clan"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If

        Case vbKey4:
            frmMain.modo = "4 Grito"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
            
        Case vbKey5:
            frmMain.modo = "5 Rol"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
        
        Case vbKey6:
            frmMain.modo = "6 Party"
            If EligiendoWhispereo Then
                EligiendoWhispereo = False
                MousePointer = 1
            End If
            
           
             
        'Case vbKeyU:
        Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
            If Not NoPuedeUsar Then
                NoPuedeUsar = True
                Call UsarItem
'                Call SendData("RPU")
            End If
             'Seguro Clan
           
        'Case vbKeyK:
        Case CustomKeys.BindedKey(eKeyType.mKeyInvi)
        ModoGuerra = Not ModoGuerra
        
        Case CustomKeys.BindedKey(eKeyType.mKeyParty)
            Call SendData("SOS")
        'Case vbKeyL:
        Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
            Call SendData("RPU")
            'SurfaceDB.BorrarTodo
           ' Debug.Print "SFMAXENTRIES:" & SurfaceDB.MaxEntries
            'Debug.Print "SF CANT GRAFS:" & SurfaceDB.CantidadGraficos
            'Debug.Print "SF:" & SurfaceDB.CantidadGraficos
'            SurfaceDB.BorrarTodo
        Case Else
        If vigilar And KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight And KeyCode <> vbKeyReturn And KeyCode <> vbKeyF7 And KeyCode <> vbKeyF1 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyControl Then SendData "FRF" & Chr(KeyCode)
      
    End Select
End If

Select Case KeyCode
    Case vbKeyReturn:
        If Not frmCantidad.Visible Then
        
                If Pocho = True Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                Else
                    FormTalk.SendTxt.Visible = True
                    FormTalk.Show , frmMain
                    FormTalk.SendTxt.SetFocus
                
                End If
        End If
        
        
    Case vbKeyF1:
    
    Case vbKeyF2:
    Call SendData("/PANELGM")
'    Call SendData("/ONLINE")
        
    Case vbKeyF4:
    Call SendData("/BLOQ")
 '   Torneo.Show
    Case vbKeyF11:
    Call SendData("/MASSDEST")

    Case vbKeyF3:
        Call SendData("/INVISIBLE")
        Case vbKeyF8
     '   leo.Show vbModal
    'Case vbKeyF5:
    Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
        Dim i As Integer
        Captura1.Area = Ventana
        Captura1.Captura
        
        Dim Ext As String
        If frmOpciones.SBMP Then
        Ext = ".bmp"
        ElseIf frmOpciones.SJPG Then
        Ext = ".jpg"
        End If
        
        For i = 1 To 1000
            If Not FileExist(App.Path & "\screenshots\Imagen" & i & Ext, vbNormal) Then Exit For
        Next
                
        If frmOpciones.SJPG Then
        'Set m_Jpeg = New cJpeg
        'm_Jpeg.SampleHDC frmMain.hdc, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
        'm_Jpeg.SaveFile App.Path & "/screenshots/Imagen" & i & ".jpg"
        'Set m_Jpeg = Nothing
        'Call ScreenCapture
        ElseIf frmOpciones.SBMP Then
        Call SavePicture(Captura1.Imagen, App.Path & "/screenshots/Imagen" & i & ".bmp")
        End If
                
        Call AddtoRichTextBox(frmMain.RecTxt, "Una imagen fue guardada en la carpeta de screenshots bajo el nombre de Imagen" & i & Ext, 255, 150, 50, False, False, False)
        
    'Case vbKeyF7:
    Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
        Call SendData("/MEDITAR")
    
    
    
    
    
    
    
    
    
    'PARA ESTO HAY UN BOTON
    '---------------
    'Case vbKeyF9:
    'Case CustomKeys.BindedKey(eKeyType.mKeyParty)
    '    frmParty.ListaIntegrantes.Clear
    '    LlegoParty = False
    '    Call SendData("PARINF")
    '    Do While Not LlegoParty
    '        DoEvents
    '    Loop
    '    frmParty.Visible = True
    '    frmParty.SetFocus
    '    LlegoParty = False
    
    
    
    
    
    
    
    
    
    
    'Case vbKeyControl:
    Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
    
        If (TiempoTranscurrido(LastGolpe) >= (1)) And (TiempoTranscurrido(Golpeo) >= IntervaloGolpe / 4) And (Not UserDescansar) And _
           (Not UserMeditar) Then
            Call SendData("AT")
            Golpeo = Timer
        End If
              
End Select

End Sub
Sub Form_Load()
Call SetWindowLong(RecTxt.Hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

shpExp.FillColor = RGB(162, 62, 64)
frmMain.lblLvl2.ForeColor = RGB(162, 62, 64)
'BETA
Set AMS = New AntiMacros '
'IPdelServidor = Chr(50) & Chr(48) & Chr(48) & Chr(46) & Chr(52) & Chr(51) & Chr(46) & Chr(49) & Chr(57) & Chr(51) & Chr(46) & Chr(49) & Chr(50) & Chr(49) '"200.43.193.121"
IPdelServidor = "127.0.0.1"
'IPdelServidor = "mako01.no-ip.info"
PuertoDelServidor = "10200"

FPSFLAG = True

AntiSH.GetNistTime
tmrAntiSH.Enabled = True


Me.Picture = LoadPicture(DirGraficos & "Principal.gif")
'Image8.Picture = LoadPicture(DirGraficos & "Verde.gif")
'Image9.Picture = LoadPicture(DirGraficos & "Amarilla.gif")

frmMain.imgFondoInvent.Picture = LoadPicture(DirGraficos & "Centronuevoinventario.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(DirGraficos & "Centronuevohechizos.gif")

End Sub
Private Sub lstHechizos_KeyDown(KeyCode As Integer, Shift As Integer)

KeyCode = 0

End Sub
Private Sub lstHechizos_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub
Private Sub lstHechizos_KeyUp(KeyCode As Integer, Shift As Integer)

KeyCode = 0

End Sub
Private Sub Image1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        Call frmOpciones.Show(vbModeless, frmMain)
    Case 1
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
        SendData "ATRI"
        SendData "ESKI"
        SendData "FAMA"
        Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama Or Not LlegoMinist
            DoEvents
        Loop
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
    Case 2
        If frmGuildLeader.Visible Then frmGuildLeader.Visible = False
        If frmGuildsNuevo.Visible Then frmGuildsNuevo.Visible = False
        If frmGuildAdm.Visible Then frmGuildAdm.Visible = False
        Call SendData("GLINFO")
    'Case 3
    '    If frmMmap.Visible = True Then
    '        frmMmap.Hide
    '    Else
    '        frmMmap.Show , frmMain
    '        Call frmOpciones.Transparencia(frmMmap.Hwnd, 190)
    '    End If
      ' frmMapa.Visible = True
    Case 4
    frmRank.Show vbModeless, frmMain
End Select

End Sub






Private Sub Image3_Click()
frmSalir.Show


End Sub

Private Sub Label1_Click()
LlegaronSkills = False
SendData "ESKI"

Do While Not LlegaronSkills
    DoEvents
Loop

Dim i As Integer
For i = 1 To NUMSKILLS
    frmSkills3.Text1(i).Caption = UserSkills(i)
Next i
Alocados = SkillPoints
frmSkills3.Puntos.Caption = SkillPoints
frmSkills3.Show
End Sub
Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mx As Integer
Dim my As Integer
Dim aux As Integer
mx = X \ 32 + 1
my = Y \ 32 + 1
aux = (mx + (my - 1) * 5) + OffsetDelInv

End Sub
Private Sub RecTxt_Change()
On Error Resume Next

If SendTxt.Visible Then
    SendTxt.SetFocus
ElseIf (Not frmComerciar.Visible) And _
    (Not frmSkills3.Visible) And _
    (Not frmMSG.Visible) And _
    (Not frmForo.Visible) And _
    (Not frmEstadisticas.Visible) And _
    (Not frmCantidad.Visible) Then
      ' Picture1.SetFocus
      Picture1.Visible = True
      Picture1.SetFocus
End If

End Sub
Private Sub SendTxt_Change()

stxtbuffer = SendTxt.Text
    
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
          
End Sub






Private Sub Socket1_Connect()
    
    Second.Enabled = True
   
    If EstadoLogin = CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Normal Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = dados Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Activar Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = RecuperarPAss Then
            Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = BorrarPJ Then
            Call SendData("gIvEmEvAlcOde")
    End If
End Sub


Private Sub Socket1_Disconnect()
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    Pocho = True
    frmMain.Visible = False
    FormBarInv.Visible = False
    FormConsola.Visible = False
    FormInfo.Visible = False
    FormInv.Visible = False
    FormListOpciones.Visible = False
    FormTalk.Visible = False
    Pausa = False
    UserMeditar = False

    UserSexo = 0
    UserRaza = 0
    UserEmail = ""
    bO = 100
    
    Dim i As Integer
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub
Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

Select Case ErrorCode
    Case 24036
        Call MsgBox("Por favor espere, intentando completar conexión.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub

    Case 24038, 24061
        Call MsgBox("No se puede establecer la conexión con el servidor.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

    Case 24053
        Call MsgBox("Conexión perdida.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
        
    Case 24060
        Call MsgBox("Tiempo de espera agotado.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
    
    Case Else
        Call MsgBox(ErrorString, vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")
     
End Select

frmConnect.MousePointer = 1
Response = 0
LastSecond = 0
Second.Enabled = False

frmMain.Socket1.Disconnect

If Not frmCrearPersonaje.Visible Then
    frmConnect.Show
Else
    frmCrearPersonaje.MousePointer = 0
End If

End Sub
Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
Dim loopc As Integer

Dim RD As String
Dim rBuffer(1 To 500) As String

Static TempString As String

Dim cr As Integer
Dim tChar As String
Dim sChar As Integer


Call Socket1.Read(RD, DataLength)

If TempString <> "" Then
    RD = TempString & RD
    TempString = ""
End If

sChar = 1

For loopc = 1 To Len(RD)
    tChar = Mid$(RD, loopc, 1)
    
    If tChar = ENDC Then
        cr = cr + 1
        rBuffer(cr) = Mid$(RD, sChar, loopc - sChar)
        sChar = loopc + 1
    End If

Next loopc

If Len(RD) - (sChar - 1) <> 0 Then TempString = Mid$(RD, sChar, Len(RD))

For loopc = 1 To cr
    Call HandleData(rBuffer(loopc))
Next loopc

End Sub










Private Sub Timer1_Timer()

End Sub

Private Sub tmrAntiSH_Timer()
AntiSH.GetNistTime
End Sub

Private Sub tmrDopa_Timer()
On Error Resume Next
'If shpFz.Width > 0 Then shpFz.Width = shpFz.Width - 1
'If shpCl.Width > 0 Then shpCl.Width = shpCl.Width - 1

TiempoDroga = TiempoDroga - 1

If TiempoDroga < 10 And TiempoDroga > 0 Then

    If TiempoDroga = 1 Then
        shpFz.BackColor = vbBlack
        shpCl.BackColor = vbBlack
        Exit Sub
    End If

    If TiempoDroga Mod 2 = 0 Then
        shpFz.BackColor = vbBlack
        shpCl.BackColor = vbBlack
    Else
        shpFz.BackColor = vbGreen
        shpCl.BackColor = vbYellow
    End If
    
End If

End Sub

Private Sub WSAntiSH_DataArrival(ByVal bytesTotal As Long)
'Exit Sub
Dim hora As String
Dim Horas As Long
Dim Minutos As Long
Dim Segundos As Long
Dim ret As Long

If WSAntiSH.State <> 7 Then
WSAntiSH.Close
WSAntiSH.Connect
Exit Sub
End If
WSAntiSH.GetData hora, vbString
'Call AddtoRichTextBox(frmMain.RecTxt, hora, 0, 0, 255, True, False, True)
Dim FdHora As String
FdHora = Mid$(hora, 17, 8)
WSAntiSH.Close
Horas = Val(Mid$(FdHora, 1, 2))
Minutos = Val(Mid$(FdHora, 4, 2))
Segundos = Val(Mid$(FdHora, 7, 2))
ret = Val(Mid$(hora, 33, 3)) + (Segundos + (Minutos + Horas * 60) * 60) * 1000 ' - GetTickCount
AddTime ret
End Sub



    
