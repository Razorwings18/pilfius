VERSION 5.00
Begin VB.UserControl LiveButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Image img_button 
      Height          =   195
      Left            =   0
      MouseIcon       =   "LiveButton.ctx":0000
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   2790
   End
End
Attribute VB_Name = "LiveButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event MouseLeave()
Public Event MouseHover()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()

Private oPictureOver As StdPicture
Private tPicture As StdPicture
Private vIsOver As Boolean

Private DidRevert As Boolean

Dim WithEvents MyTrak As clsTrackInfo
Attribute MyTrak.VB_VarHelpID = -1

Private Sub img_button_Click()
    RaiseEvent Click
End Sub

Private Sub MyTrak_MouseHover()
    RaiseEvent MouseHover
End Sub

Private Sub DoButtonLeave()
    If Not tPicture Is Nothing Then
        vIsOver = False
        Set img_button.Picture = tPicture
        Set tPicture = Nothing
    End If
End Sub

Private Sub MyTrak_MouseLeave()
    DoButtonLeave
    RaiseEvent MouseLeave
End Sub

Public Property Get HoverTime() As Long
HoverTime = MyTrak.HoverTime
End Property

Public Property Let HoverTime(newHoverTime As Long)
MyTrak.HoverTime = newHoverTime
PropertyChanged "HoverTime"
End Property

Private Sub UserControl_GotFocus()
    If Not DidRevert Then
        DoButtonOver
    Else
        DidRevert = False
    End If
End Sub

Private Sub UserControl_InitProperties()
Set MyTrak = New clsTrackInfo
MyTrak.HoverTime = 400

vIsOver = False
DidRevert = False
End Sub

Private Sub DoButtonOver()
    If Not oPictureOver Is Nothing And Not vIsOver Then
        vIsOver = True
        Set tPicture = img_button.Picture
        Set img_button.Picture = oPictureOver
    End If
End Sub

Private Sub img_button_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoButtonOver
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_LostFocus()
    DoButtonLeave
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set MyTrak = New clsTrackInfo
MyTrak.hwnd = UserControl.hwnd

MyTrak.HoverTime = PropBag.ReadProperty("HoverTime", 400)
UserControl.BackColor = PropBag.ReadProperty("BackColor", RGB(200, 200, 200))
Set img_button.Picture = PropBag.ReadProperty("Picture", Nothing)
Set oPictureOver = PropBag.ReadProperty("PictureOver", Nothing)

If Ambient.UserMode Then
StartTrack MyTrak
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "HoverTime", MyTrak.HoverTime, 400
PropBag.WriteProperty "Picture", img_button.Picture, Nothing
PropBag.WriteProperty "PictureOver", oPictureOver, Nothing
PropBag.WriteProperty "BackColor", UserControl.BackColor, RGB(200, 200, 200)
End Sub

Private Sub UserControl_Terminate()
EndTrack MyTrak
Set MyTrak = Nothing
End Sub

Public Property Get Picture() As StdPicture
    Set Picture = img_button.Picture
End Property

Public Property Set Picture(pPicture As StdPicture)
    Set img_button.Picture = pPicture
    
    UserControl.Width = img_button.Width
    UserControl.Height = img_button.Height
    
    PropertyChanged "Picture"
End Property

Public Property Get PictureOver() As StdPicture
    Set PictureOver = oPictureOver
End Property

Public Property Set PictureOver(pPictureOver As StdPicture)
    Set oPictureOver = pPictureOver
    
    PropertyChanged "PictureOver"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(pColor As OLE_COLOR)
    UserControl.BackColor = pColor
    
    PropertyChanged "BackColor"
End Property

Public Sub RevertPicture()
    If vIsOver Then
        DoButtonLeave
        DidRevert = True
        RaiseEvent MouseLeave
    End If
End Sub
