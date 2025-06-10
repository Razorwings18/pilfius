VERSION 5.00
Begin VB.Form frm_addcommand 
   BackColor       =   &H0078FFEA&
   BorderStyle     =   0  'None
   Caption         =   "Add command"
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PiLfIuS.LiveButton cmd_cancel 
      Height          =   420
      Left            =   4770
      TabIndex        =   3
      Top             =   1605
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   741
      Picture         =   "frm_addcommand.frx":0000
      PictureOver     =   "frm_addcommand.frx":062F
      BackColor       =   7929834
   End
   Begin PiLfIuS.LiveButton cmd_add 
      Height          =   420
      Left            =   4215
      TabIndex        =   2
      Top             =   1605
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   741
      Picture         =   "frm_addcommand.frx":0C4F
      PictureOver     =   "frm_addcommand.frx":1272
      BackColor       =   7929834
   End
   Begin VB.ComboBox cmb_group 
      Appearance      =   0  'Flat
      BackColor       =   &H0042F0D6&
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
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1155
      Width           =   3105
   End
   Begin VB.TextBox txt_command 
      Appearance      =   0  'Flat
      BackColor       =   &H0042F0D6&
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
      Left            =   1710
      TabIndex        =   0
      Top             =   645
      Width           =   2850
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0033C9B4&
      Height          =   2085
      Left            =   0
      Top             =   0
      Width           =   5370
   End
   Begin VB.Image Image3 
      Height          =   150
      Left            =   1065
      Picture         =   "frm_addcommand.frx":188D
      Top             =   1230
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   120
      Left            =   555
      Picture         =   "frm_addcommand.frx":19B2
      Top             =   735
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   150
      Left            =   375
      Picture         =   "frm_addcommand.frx":1B14
      Top             =   285
      Width           =   3870
   End
End
Attribute VB_Name = "frm_addcommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancelled As Boolean
Public Command As String
Public GroupNr As Integer

Private Sub cmb_group_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmd_add_Click
    End If
    KeyHandlerGeneral KeyCode, Shift
End Sub

Private Sub cmd_add_Click()
    Dim errmsg As String
    
    errmsg = ""
    If Len(Trim(txt_command.Text)) = 0 Then errmsg = errmsg & "- Please, enter the spoken command to recognize"
    
    If Len(errmsg) = 0 Then
        Cancelled = False
        Command = txt_command.Text
        txt_command.Text = ""
        
        Me.Hide
    Else
        MsgBox errmsg, vbOKOnly + vbExclamation, "MISSING INPUT"
    End If
End Sub

Private Sub cmd_cancel_Click()
    Cancelled = True
    Command = ""
    txt_command.Text = ""
    
    Me.Hide
End Sub

Private Sub Form_Load()
    LoadGroups
    
    Do While cmb_group.ItemData(cmb_group.ListIndex) <> GroupNr
        cmb_group.ListIndex = cmb_group.ListIndex + 1
    Loop
    
    If Len(Command) > 0 Then
        txt_command.Text = Command
        
        'cmd_add.Caption = "Modify command"
    Else
        'cmd_add.Caption = "Add command"
    End If
End Sub

Private Sub LoadGroups()
    cmb_group.AddItem "--- None ---"
    cmb_group.ItemData(0) = 0
    
    For i = 1 To fra_main.oCommand.cGroups.Count
        cmb_group.AddItem fra_main.oCommand.cGroups(i)
        cmb_group.ItemData(cmb_group.NewIndex) = i
    Next
    cmb_group.ListIndex = 0
End Sub

Private Sub KeyHandlerGeneral(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27: cmd_cancel_Click
    End Select
End Sub

Private Sub txt_command_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmd_add_Click
    End If
    KeyHandlerGeneral KeyCode, Shift
End Sub
