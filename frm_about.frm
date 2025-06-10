VERSION 5.00
Begin VB.Form frm_about 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   13590
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
      Height          =   4245
      Left            =   8220
      Picture         =   "frm_about.frx":0000
      ScaleHeight     =   4185
      ScaleWidth      =   5310
      TabIndex        =   0
      Top             =   -15
      Width           =   5370
   End
   Begin VB.Image cmd_close 
      Height          =   150
      Left            =   3570
      MouseIcon       =   "frm_about.frx":2856
      MousePointer    =   99  'Custom
      Picture         =   "frm_about.frx":29A8
      Top             =   75
      Width           =   1455
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   5085
      Picture         =   "frm_about.frx":2D5C
      Top             =   3855
      Width           =   2040
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   2025
      Picture         =   "frm_about.frx":3164
      Top             =   3525
      Width           =   2565
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   1875
      Picture         =   "frm_about.frx":36B7
      Top             =   2610
      Width           =   2610
   End
   Begin VB.Image img_web 
      Height          =   375
      Left            =   1575
      MouseIcon       =   "frm_about.frx":3C22
      MousePointer    =   99  'Custom
      Picture         =   "frm_about.frx":3D74
      Top             =   2115
      Width           =   2595
   End
   Begin VB.Image Image2 
      Height          =   780
      Left            =   1170
      Picture         =   "frm_about.frx":4290
      Top             =   1020
      Width           =   2760
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOW = 5

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

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
   Me.Width = 7125 'picBackground.ScaleWidth
   Me.Height = picBackground.ScaleHeight
   Set Me.Picture = picBackground.Picture
   
   Dim transColor As Long
   transColor = &H8000FF
   Me.BackColor = transColor
   
   Dim lStyle As Long
   lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
   lStyle = lStyle Or WS_EX_LAYERED
   SetWindowLong hWnd, GWL_EXSTYLE, lStyle
      
   SetLayeredWindowAttributes Me.hWnd, transColor, 220, LWA_COLORKEY Or LWA_ALPHA
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

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub img_web_Click()
    Call ShellExecute(Me.hWnd, "open", "http://www.pilfius.com.ar/", vbNullString, vbNullString, SW_SHOW)
End Sub
