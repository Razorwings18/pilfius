VERSION 5.00
Begin VB.Form frm_changeactivation 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_detect2 
      BackColor       =   &H00A5C944&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   11865
      TabIndex        =   11
      Top             =   2835
      Visible         =   0   'False
      Width           =   1950
      Begin PiLfIuS.LiveButton cmd_joy_cancel 
         Height          =   300
         Left            =   375
         TabIndex        =   14
         Top             =   765
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Picture         =   "frm_changeactivation.frx":0000
         PictureOver     =   "frm_changeactivation.frx":032A
         BackColor       =   10864964
      End
      Begin PiLfIuS.LiveButton cmd_joy_ok 
         Height          =   300
         Left            =   30
         TabIndex        =   13
         Top             =   765
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Picture         =   "frm_changeactivation.frx":0654
         PictureOver     =   "frm_changeactivation.frx":0978
         BackColor       =   10864964
      End
   End
   Begin VB.Frame fra_detect 
      BackColor       =   &H00A5C944&
      BorderStyle     =   0  'None
      Height          =   1650
      Left            =   8820
      TabIndex        =   10
      Top             =   2835
      Visible         =   0   'False
      Width           =   3030
      Begin VB.Label lbl_joydetect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOY1 BTN 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1380
         TabIndex        =   12
         Top             =   765
         Width           =   885
      End
      Begin VB.Image Image7 
         Height          =   150
         Left            =   300
         Picture         =   "frm_changeactivation.frx":0C9E
         Top             =   810
         Width           =   1005
      End
      Begin VB.Image Image6 
         Height          =   300
         Left            =   300
         Picture         =   "frm_changeactivation.frx":0E05
         Top             =   255
         Width           =   2280
      End
   End
   Begin VB.Frame fra_key 
      BackColor       =   &H00A5C944&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1050
      TabIndex        =   7
      Top             =   735
      Width           =   4335
      Begin VB.ComboBox cmb_special 
         Appearance      =   0  'Flat
         BackColor       =   &H00BAE070&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   1650
      End
      Begin VB.TextBox txt_keyactivation 
         Appearance      =   0  'Flat
         BackColor       =   &H00BAE070&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   15
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   150
         Left            =   1650
         Picture         =   "frm_changeactivation.frx":1265
         Top             =   120
         Width           =   915
      End
   End
   Begin VB.ComboBox cmb_mouse 
      Appearance      =   0  'Flat
      BackColor       =   &H00BAE070&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1695
      Width           =   1650
   End
   Begin VB.OptionButton opt_activation 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2CC4C&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   435
      TabIndex        =   3
      Top             =   1785
      Width           =   225
   End
   Begin VB.OptionButton opt_activation 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2CC4C&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   435
      TabIndex        =   2
      Top             =   1320
      Width           =   225
   End
   Begin VB.OptionButton opt_activation 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2CC4C&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   435
      TabIndex        =   1
      Top             =   825
      Width           =   225
   End
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
      Height          =   2745
      Left            =   8820
      Picture         =   "frm_changeactivation.frx":13AD
      ScaleHeight     =   2685
      ScaleWidth      =   5565
      TabIndex        =   0
      Top             =   0
      Width           =   5625
   End
   Begin PiLfIuS.LiveButton cmd_ok 
      Height          =   435
      Left            =   4290
      TabIndex        =   5
      Top             =   2070
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   767
      Picture         =   "frm_changeactivation.frx":3316
      PictureOver     =   "frm_changeactivation.frx":375B
      BackColor       =   4000906
   End
   Begin PiLfIuS.LiveButton cmd_cancel 
      Height          =   435
      Left            =   4815
      TabIndex        =   6
      Top             =   2070
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   767
      Picture         =   "frm_changeactivation.frx":3B8B
      PictureOver     =   "frm_changeactivation.frx":3FD5
      BackColor       =   4000906
   End
   Begin VB.Frame fra_joy 
      BackColor       =   &H00A5C944&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2310
      TabIndex        =   15
      Top             =   1290
      Width           =   2985
      Begin PiLfIuS.LiveButton cmd_joyselect 
         Height          =   240
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   423
         Picture         =   "frm_changeactivation.frx":442E
         PictureOver     =   "frm_changeactivation.frx":48B1
         BackColor       =   10669132
      End
      Begin VB.Label lbl_joy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOY1 BTN 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   870
         TabIndex        =   17
         Top             =   15
         Width           =   885
      End
   End
   Begin VB.Image Image5 
      Height          =   150
      Left            =   705
      Picture         =   "frm_changeactivation.frx":4CB6
      Top             =   1815
      Width           =   825
   End
   Begin VB.Image Image4 
      Height          =   150
      Left            =   705
      Picture         =   "frm_changeactivation.frx":4DEC
      Top             =   1350
      Width           =   1515
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   705
      Picture         =   "frm_changeactivation.frx":4FC7
      Top             =   855
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   150
      Left            =   435
      Picture         =   "frm_changeactivation.frx":5078
      Top             =   480
      Width           =   2160
   End
End
Attribute VB_Name = "frm_changeactivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------ TRANSPARENT WINDOW STUFF -----------------------------------------------
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
'--------------------------------------- END TRANSPARENT WINDOW STUFF -----------------------------------------------

Public Cancelled As Boolean
Public selectedAction As String
Private stopJoyDetection As Boolean
Private lastOpt As Integer

Private Sub cmb_special_Click()
    txt_keyactivation.text = ""
End Sub

Private Sub cmd_cancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub cmd_joy_cancel_Click()
    stopJoyDetection = True
    fra_detect.Visible = False
    fra_detect2.Visible = False
    lbl_joydetect.Caption = ""
End Sub

Private Sub cmd_joy_ok_Click()
    stopJoyDetection = True
    If Len(lbl_joydetect.Caption) > 0 Then lbl_joy.Caption = lbl_joydetect.Caption
    
    fra_detect.Visible = False
    fra_detect2.Visible = False
    lbl_joydetect.Caption = ""
End Sub

Private Sub cmd_joyselect_Click()
    Dim installedJoys As Integer
    Dim pressedBtns As Integer
    Dim lastpressedBtns As Integer
    Dim resetState As Boolean
    Dim joyCaption As String
    Dim initialPOVstate As Variant
    
    Dim oJoystick As cls_joystick
    
    Set oJoystick = New cls_joystick
    
    fra_detect.Visible = True
    fra_detect2.Visible = True
    
    installedJoys = oJoystick.countInstalledJoys
    If installedJoys > 0 Then
        initialPOVstate = oJoystick.getJoysInitialPOVState(installedJoys)
        
        stopJoyDetection = False
        resetState = False
        Do While Not stopJoyDetection
            joyCaption = oJoystick.DoJoyDetection(installedJoys, initialPOVstate, pressedBtns)
            If Len(joyCaption) = 0 Then resetState = True
            
            If Len(joyCaption) > 0 Then
                If resetState Or (Not resetState And (pressedBtns >= lastpressedBtns)) Then
                    lbl_joydetect.Caption = joyCaption
                    
                    lastpressedBtns = pressedBtns
                    resetState = False
                End If
            End If
            
            DoEvents
        Loop
    End If
End Sub

Private Sub cmd_ok_Click()
    Dim msgerror As String
    
    selectedAction = ""
    Select Case True
        Case opt_activation(0):
                                If Len(txt_keyactivation.text) > 0 Then
                                    selectedAction = txt_keyactivation.text
                                ElseIf cmb_special.ListIndex > 0 Then
                                    selectedAction = cmb_special.List(cmb_special.ListIndex)
                                End If
        Case opt_activation(1): selectedAction = lbl_joy.Caption
        Case opt_activation(2): If cmb_mouse.ListIndex > 0 Then selectedAction = cmb_mouse.List(cmb_mouse.ListIndex)
    End Select
    
    msgerror = ""
    If Len(selectedAction) = 0 Then msgerror = msgerror & "- No input selected!"
    
    If Len(msgerror) = 0 Then
        Cancelled = False
        Me.Hide
    Else
        MsgBox msgerror, vbExclamation + vbOKOnly, "CANNOT ASSIGN INPUT"
    End If
End Sub

Private Sub Form_Activate()
    Me.Left = (mdi_main.Left + (mdi_main.Width / 2)) - (Me.Width / 2) - 800
    Me.Top = (mdi_main.Top + (mdi_main.Height / 2)) - (Me.Height / 2)
    
    resetForm
    If Len(selectedAction) > 0 Then preselectActivation
End Sub

Private Sub preselectActivation()
    If Mid(selectedAction, 1, 3) = "JOY" Then
        opt_activation(1).Value = True
        lbl_joy.Caption = selectedAction
    ElseIf Mid(selectedAction, 1, 5) = "MOUSE" Then
        opt_activation(2).Value = True
        
        Do While cmb_mouse.List(cmb_mouse.ListIndex) <> selectedAction
            cmb_mouse.ListIndex = cmb_mouse.ListIndex + 1
        Loop
    Else
        opt_activation(0).Value = True
        txt_keyactivation.text = selectedAction
    End If
End Sub

Private Sub resetForm()
    Dim obj As Object
    
    For Each obj In opt_activation
        obj.Value = False
    Next
    txt_keyactivation.text = ""
    cmb_special.ListIndex = 0
    cmb_mouse.ListIndex = 0
    lbl_joy.Caption = ""
    lbl_joydetect.Caption = ""
    
    fra_detect.Visible = False
    fra_detect2.Visible = False
    
    Cancelled = False
    
    lastOpt = -1
    hideOptionStuff
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
      
   SetLayeredWindowAttributes Me.hWnd, transColor, 240, LWA_COLORKEY Or LWA_ALPHA
   
   fra_detect.Top = 475
   fra_detect.Left = 390
   fra_detect2.Top = fra_detect.Top
   fra_detect2.Left = fra_detect.Left + fra_detect.Width
   
   fra_main.FillSpecialKeys cmb_special
   fra_main.FillMouseJoy cmb_mouse, False, False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessageLong Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub hideOptionStuff()
    fra_key.Visible = False
    fra_joy.Visible = False
    cmb_mouse.Visible = False
End Sub

Private Sub opt_activation_Click(Index As Integer)
    If lastOpt <> Index Then
        hideOptionStuff
        
        Select Case True
            Case opt_activation(0).Value: fra_key.Visible = True
                                            lastOpt = 0
            Case opt_activation(1).Value: fra_joy.Visible = True
                                            lastOpt = 1
            Case opt_activation(2).Value: cmb_mouse.Visible = True
                                            lastOpt = 2
        End Select
    End If
End Sub

Private Sub txt_keyactivation_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim comboName As String
    
    cmb_special.ListIndex = 0
    
    comboName = fra_main.getComboName(KeyCode, Shift)
    
    txt_keyactivation.text = comboName
End Sub
