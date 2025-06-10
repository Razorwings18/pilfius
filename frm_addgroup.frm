VERSION 5.00
Begin VB.Form frm_addgroup 
   BackColor       =   &H0078FFEA&
   BorderStyle     =   0  'None
   Caption         =   "Add group"
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_group 
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
      Left            =   1395
      TabIndex        =   0
      Top             =   645
      Width           =   3135
   End
   Begin PiLfIuS.LiveButton cmd_cancel 
      Height          =   420
      Left            =   4185
      TabIndex        =   2
      Top             =   1155
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   741
      Picture         =   "frm_addgroup.frx":0000
      PictureOver     =   "frm_addgroup.frx":062F
      BackColor       =   7929834
   End
   Begin PiLfIuS.LiveButton cmd_add 
      Height          =   420
      Left            =   3630
      TabIndex        =   1
      Top             =   1155
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   741
      Picture         =   "frm_addgroup.frx":0C4F
      PictureOver     =   "frm_addgroup.frx":1272
      BackColor       =   7929834
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0033C9B4&
      Height          =   1710
      Left            =   0
      Top             =   0
      Width           =   4995
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   465
      Picture         =   "frm_addgroup.frx":188D
      Top             =   750
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   150
      Left            =   465
      Picture         =   "frm_addgroup.frx":19EF
      Top             =   300
      Width           =   1545
   End
End
Attribute VB_Name = "frm_addgroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cancelled As Boolean
Public Group As String

Private Sub cmd_add_Click()
    Dim errmsg As String
    
    errmsg = ""
    If Len(Trim(txt_group.Text)) = 0 Then errmsg = errmsg & "- Please, enter the group's name"
    
    If Len(errmsg) = 0 Then
        Cancelled = False
        Group = txt_group.Text
        
        Me.Hide
    Else
        MsgBox errmsg, vbOKOnly + vbExclamation, "MISSING INPUT"
    End If
End Sub

Private Sub cmd_cancel_Click()
    Group = ""
    
    Cancelled = True
    Me.Hide
End Sub

Private Sub Form_Load()
    If Len(Group) > 0 Then txt_group.Text = Group
End Sub

Private Sub txt_group_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmd_add_Click
    End If
    KeyHandlerGeneral KeyCode, Shift
End Sub

Private Sub KeyHandlerGeneral(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27: cmd_cancel_Click
    End Select
End Sub

