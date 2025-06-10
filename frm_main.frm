VERSION 5.00
Begin VB.Form frm_main 
   BackColor       =   &H0074FFEA&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MouseIcon       =   "frm_main.frx":0000
   Picture         =   "frm_main.frx":0152
   ScaleHeight     =   8505
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr_slidehelp 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   5985
      Top             =   60
   End
   Begin VB.Frame fra_fake 
      BackColor       =   &H009EFEEE&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   510
      Left            =   300
      TabIndex        =   2
      Top             =   0
      Width           =   3750
      Begin VB.Image Image2 
         Height          =   120
         Left            =   255
         Picture         =   "frm_main.frx":186BD
         Top             =   285
         Width           =   1290
      End
   End
   Begin VB.Frame fra_menu 
      BackColor       =   &H0077F2DF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   675
      TabIndex        =   1
      Top             =   705
      Width           =   2895
      Begin PiLfIuS.LiveButton cmd_SREsetup 
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   675
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   344
         Picture         =   "frm_main.frx":1884F
         BackColor       =   7860959
      End
      Begin PiLfIuS.LiveButton img_load 
         Height          =   180
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   318
         Picture         =   "frm_main.frx":18B7A
         BackColor       =   7860959
      End
      Begin PiLfIuS.LiveButton img_new 
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   330
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   318
         Picture         =   "frm_main.frx":18E21
         BackColor       =   7860959
      End
   End
   Begin VB.Timer tmr_beg 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   5520
      Top             =   60
   End
   Begin VB.Image img_about 
      Height          =   375
      Left            =   9585
      MouseIcon       =   "frm_main.frx":1916A
      MousePointer    =   99  'Custom
      Picture         =   "frm_main.frx":192BC
      Top             =   7740
      Width           =   1710
   End
   Begin VB.Line Line3 
      X1              =   8700
      X2              =   285
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label lbl_version 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver 0.6"
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
      Left            =   10110
      TabIndex        =   0
      Top             =   855
      Width           =   540
   End
   Begin VB.Line Line2 
      X1              =   9510
      X2              =   11430
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line1 
      X1              =   11415
      X2              =   11415
      Y1              =   1470
      Y2              =   7920
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   5355
      Picture         =   "frm_main.frx":19D70
      Top             =   7410
      Width           =   4155
   End
   Begin VB.Image img_menu 
      Appearance      =   0  'Flat
      Height          =   1680
      Left            =   315
      Picture         =   "frm_main.frx":1B457
      Top             =   525
      Width           =   3570
   End
   Begin VB.Image img_txt_helper 
      Height          =   510
      Left            =   945
      Picture         =   "frm_main.frx":1C35A
      Top             =   720
      Width           =   2580
   End
   Begin VB.Image img_helper 
      Height          =   855
      Left            =   585
      Picture         =   "frm_main.frx":1CA28
      Top             =   570
      Width           =   3390
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum e_HelpShowMode
    hsm_Show = 0
    hsm_Hide = 1
End Enum

Dim HelpShowMode As e_HelpShowMode
Dim load_isIn As Boolean
Dim showSetup As Boolean
Dim oSpeech As cls_speech


Private Sub cmd_SREsetup_Click()
    runSREsetup Me.hWnd
End Sub

Private Sub Form_Activate()
    mdi_main.commandlineParse
End Sub

Private Sub Form_Load()
    Left = 0
    Top = 0
    
    img_menu.Top = img_menu.Top - 1200
    img_helper.Top = img_helper.Top - 1200
    img_txt_helper.Top = img_txt_helper.Top - 1200
    fra_menu.Top = fra_menu.Top - 1200
    
    lbl_version.Caption = mdi_main.vVersionNumber
    
    tmr_beg.Enabled = True
    load_isIn = False
    
    showSetup = True
End Sub

Private Sub img_about_Click()
    frm_about.Show
End Sub

Private Sub img_load_Click()
    img_load.RevertPicture
    img_load_MouseLeave
    mdi_main.mnu_file_open_Click
End Sub

Private Sub img_load_MouseLeave()
    HideHelp
End Sub

Private Sub HideHelp()
    If load_isIn Then
        HelpShowMode = hsm_Hide
        tmr_slidehelp.Enabled = True
        
        load_isIn = False
    End If
End Sub

Private Sub ShowHelp()
    If Not load_isIn Then
        HelpShowMode = hsm_Show
        tmr_slidehelp.Enabled = True
        
        load_isIn = True
    End If
End Sub

Private Sub img_load_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHelp
End Sub

Private Sub img_new_Click()
    img_new.RevertPicture
    img_new_MouseLeave
    mdi_main.mnu_file_create_Click
End Sub

Private Sub img_new_MouseLeave()
    HideHelp
End Sub

Private Sub img_new_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHelp
End Sub

Private Sub LiveButton1_Click()
    
End Sub

Private Sub LiveButton1_MouseLeave()

End Sub

Private Sub tmr_beg_Timer()
    If showSetup Then
        'Check if it's the first time PiLfIuS! is run
        isFirstTime = GetSetting("PILFIUS", "installation-specific", "runonce", 0)
        If isFirstTime = 0 Then
            frm_setup.Show
            showSetup = False
            
            SaveSetting "PILFIUS", "installation-specific", "runonce", 1
        End If
    End If
    
    img_menu.Top = img_menu.Top + 60
    img_helper.Top = img_helper.Top + 60
    img_txt_helper.Top = img_txt_helper.Top + 60
    fra_menu.Top = fra_menu.Top + 60
    
    If img_menu.Top >= 525 Then tmr_beg.Enabled = False
End Sub

Private Sub tmr_slidehelp_Timer()
    Select Case HelpShowMode
        Case 0:
                If img_helper.Left >= 3400 Then
                    tmr_slidehelp.Enabled = False
                Else
                    img_helper.Left = img_helper.Left + 160
                    img_txt_helper.Left = img_txt_helper.Left + 160
                End If
        Case 1:
                If img_helper.Left <= 585 Then
                    tmr_slidehelp.Enabled = False
                Else
                    img_helper.Left = img_helper.Left - 160
                    img_txt_helper.Left = img_txt_helper.Left - 160
                End If
    End Select
End Sub

Public Sub runSREsetup(parenthWnd As Long)
    Set oSpeech = New cls_speech
    If oSpeech.InitRecognition Then
        oSpeech.runMicTrainer parenthWnd
        oSpeech.runUserTrainer parenthWnd
        oSpeech.runProfileProperties parenthWnd
    End If
    Set oSpeech = Nothing
End Sub
