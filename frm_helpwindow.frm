VERSION 5.00
Begin VB.Form frm_helpwindow 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBackground 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   8955
      Picture         =   "frm_helpwindow.frx":0000
      ScaleHeight     =   5055
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   4980
      Width           =   8055
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   1695
      Picture         =   "frm_helpwindow.frx":2C42
      Top             =   1170
      Width           =   5670
   End
   Begin VB.Label lbl_text 
      BackStyle       =   0  'Transparent
      Caption         =   "If checked, ""use Confidence Threshold"" will verify how certain the Speech"
      ForeColor       =   &H00FFFFFF&
      Height          =   2925
      Left            =   1695
      TabIndex        =   2
      Top             =   1350
      Width           =   5865
   End
   Begin VB.Label lbl_title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What is ""use Confidence Threshold""?"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1935
      TabIndex        =   1
      Top             =   900
      Width           =   2685
   End
   Begin VB.Image Image1 
      Height          =   90
      Left            =   1695
      Picture         =   "frm_helpwindow.frx":2D3D
      Top             =   960
      Width           =   90
   End
   Begin VB.Image cmd_close 
      Height          =   150
      Left            =   6120
      MouseIcon       =   "frm_helpwindow.frx":2D95
      MousePointer    =   99  'Custom
      Picture         =   "frm_helpwindow.frx":2EE7
      Top             =   780
      Width           =   1425
   End
End
Attribute VB_Name = "frm_helpwindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

'Requires Windows 2000 or later:
Private Const WS_EX_LAYERED = &H80000
Private Type BLENDFUNCTION
   BlendOp As Byte
   BlendFlags As Byte
   SourceConstantAlpha As Byte
   AlphaFormat As Byte
End Type

Private Const AC_SRC_OVER = &H0

Private Const AC_SRC_ALPHA = &H1
Private Const AC_SRC_NO_PREMULT_ALPHA = &H1
Private Const AC_SRC_NO_ALPHA = &H2
Private Const AC_DST_NO_PREMULT_ALPHA = &H10
Private Const AC_DST_NO_ALPHA = &H20

Private Declare Function SetLayeredWindowAttributes Lib "USER32" _
   (ByVal hWnd As Long, ByVal crKey As Long, _
   ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Declare Function UpdateLayeredWindow Lib "USER32" _
   (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, _
   psize As Any, ByVal hdcSrc As Long, _
   pptSrc As Any, crKey As Long, _
   ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4

Private Declare Function RedrawWindow Lib "USER32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_FRAME = &H400
Private Const RDW_INVALIDATE = &H1


Private Sub cmd_close_click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Left = (mdi_main.Left + (mdi_main.Width / 2)) - (Me.Width / 2)
    Me.Top = (mdi_main.Top + (mdi_main.Height / 2)) - (Me.Height / 2)
End Sub

Private Sub Form_Load()
   Me.Width = picBackground.ScaleWidth
   Me.Height = picBackground.ScaleHeight
   Set Me.Picture = picBackground.Picture
   
   Dim transColor As Long
   transColor = &H8000FF
   Me.BackColor = transColor
   
   Dim lStyle As Long
   lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
   lStyle = lStyle Or WS_EX_LAYERED
   SetWindowLong hWnd, GWL_EXSTYLE, lStyle
      
   SetLayeredWindowAttributes Me.hWnd, transColor, 250, LWA_COLORKEY Or LWA_ALPHA
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessageLong Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lbl_text_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lbl_title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub
