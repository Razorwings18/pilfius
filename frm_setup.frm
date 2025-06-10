VERSION 5.00
Begin VB.Form frm_setup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frm_setup.frx":0000
   ScaleHeight     =   3705
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PiLfIuS.LiveButton cmd_skip 
      Height          =   420
      Left            =   3825
      TabIndex        =   1
      Top             =   2340
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   741
      Picture         =   "frm_setup.frx":3E8C
      PictureOver     =   "frm_setup.frx":4561
      BackColor       =   4842738
   End
   Begin PiLfIuS.LiveButton cmd_setup 
      Height          =   420
      Left            =   855
      TabIndex        =   0
      Top             =   2340
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   741
      Picture         =   "frm_setup.frx":4ADD
      PictureOver     =   "frm_setup.frx":53A9
      BackColor       =   4842738
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   975
      Picture         =   "frm_setup.frx":5AFB
      Top             =   1275
      Width           =   5310
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   2910
      Picture         =   "frm_setup.frx":6AB8
      Top             =   690
      Width           =   2730
   End
End
Attribute VB_Name = "frm_setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_setup_Click()
    frm_main.runSREsetup Me.hWnd
    MsgBox "Thank you for taking the time to set up the Speech Recognition Engine. Should you ever need to access the Speech Recognition Engine configuration again, just click on the ""Speech Engine settings"" menu item at the main screen.", vbOKOnly + vbInformation, "SETUP COMPLETE"
    
    frm_main.Enabled = True
    
    Unload Me
End Sub

Private Sub cmd_skip_Click()
    frm_main.Enabled = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    frm_main.Enabled = False
    
    Me.Left = 2700
    Me.Top = 2300
End Sub
