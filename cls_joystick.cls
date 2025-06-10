VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_joystick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------- JOYSTICK STUFF --------------------------------------------------
Private Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Private Declare Function joyGetNumDevs Lib "winmm.dll" () As Long

Const JOY_RETURNBUTTONS = &H80&
Const JOY_RETURNCENTERED = &H400&
Const JOY_RETURNPOV = &H40&
Const JOY_RETURNPOVCTS = &H200&
Const JOY_RETURNR = &H8&
Const JOY_RETURNRAWDATA = &H100&
Const JOY_RETURNU = &H10                             '  axis 5
Const JOY_RETURNV = &H20                             '  axis 6
Const JOY_RETURNX = &H1&
Const JOY_RETURNY = &H2&
Const JOY_RETURNZ = &H4&
Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)

Private Type JOYINFOEX
        dwSize As Long                '  size of structure
        dwFlags As Long               '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                '  rudder/4th axis position
        dwUpos As Long                '  5th axis position
        dwVpos As Long                '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long           '  reserved for communication between winmm driver
        dwReserved2 As Long           '  reserved for future expansion
End Type

'------------------------------------- END JOYSTICK STUFF --------------------------------------------------

Public Function countInstalledJoys() As Integer
    'Returns number of installed joysticks in this system
    
    Dim installedJoys As Integer
    Dim maxNumDevs As Integer
    Dim joyInfo As JOYINFOEX
    Dim joyIDquery As Integer
    
    maxNumDevs = joyGetNumDevs
    
    installedJoys = 0
    For joyIDquery = 0 To maxNumDevs
        joyInfo.dwSize = Len(joyInfo)
        joyInfo.dwFlags = JOY_RETURNALL
        
        errNumber = joyGetPosEx(joyIDquery, joyInfo)
        If errNumber = 0 Then installedJoys = installedJoys + 1
    Next
    
    countInstalledJoys = installedJoys
End Function

Public Function getJoysInitialPOVState(installedJoys As Integer) As Variant
    'INPUT: installedJoys (number of joysticks installed in this system)
    'OUTPUT: Array of (installedjoys) with the idle POV return value.
    
    Dim joyInfo As JOYINFOEX
    Dim joyNum As Integer
    Dim joyIDquery As Integer
    
    Dim vBytePOV As Variant
    
    If installedJoys > 0 Then
        ReDim vBytePOV(installedJoys - 1)
        For joyIDquery = 0 To installedJoys - 1
            joyInfo.dwSize = Len(joyInfo)
            joyInfo.dwFlags = JOY_RETURNALL
            
            errNumber = joyGetPosEx(joyIDquery, joyInfo)
            vBytePOV(joyIDquery) = joyInfo.dwPOV
        Next
        
        getJoysInitialPOVState = vBytePOV
    Else
        getJoysInitialPOVState = Null
    End If
End Function

Public Function DoJoyDetection(installedJoys As Integer, initialPOVstate As Variant, ByRef pressedBtns As Integer) As String
    'INPUT: installedJoys (number of joysticks installed in this system), pressedBtns (returns number of pressed buttons here)
    'OUTPUT: A string containing joystick number an buttons pressed or HAT position.
    
    Dim joyInfo As JOYINFOEX
    Dim joyNum As Integer
    Dim joyIDquery As Integer
    
    Dim errNumber As Long
    Dim byteVal As Long
    Dim lastBtnNumber As Long
    Dim bytePOV As Long
    Dim i As Integer
    Dim btnsBin As String
    Dim hadBTN As Boolean
    Dim btnString As String
    Dim hatString As String
    Dim resetState As Boolean
    
    If installedJoys > 0 Then
        For joyIDquery = 0 To installedJoys - 1
            joyInfo.dwSize = Len(joyInfo)
            joyInfo.dwFlags = JOY_RETURNALL
            
            errNumber = joyGetPosEx(joyIDquery, joyInfo)
            byteVal = joyInfo.dwButtons
            bytePOV = joyInfo.dwPOV
            
            If byteVal > 0 Then
                btnsBin = ""
                Do While byteVal > 0
                    btnsBin = btnsBin & CStr(byteVal Mod 2)
                    byteVal = byteVal \ 2
                Loop
                
                btnsBin = StrReverse(btnsBin)
                
                hadBTN = False
                btnString = ""
                For i = Len(btnsBin) To 1 Step -1
                    If Mid(btnsBin, i, 1) = "1" Then
                        If hadBTN Then btnString = btnString & "+"
                        btnString = btnString & (Len(btnsBin) - (i - 1))
                        
                        hadBTN = True
                    End If
                Next
                
                DoJoyDetection = "JOY" & joyIDquery & " BTN " & btnString
                
                pressedBtns = joyInfo.dwButtonNumber
            End If
            
            If bytePOV <> initialPOVstate(joyIDquery) Then
                hatString = bytePOV / 100
                
                DoJoyDetection = "JOY" & joyIDquery & " HAT " & hatString
            End If
        Next
    End If
End Function

Public Function joyIsButtonStatus(joyID As Integer, byteVal As Long) As Boolean
    'Returns TRUE if AT LEAST joystick buttons specified in byteVal are pressed.
    Dim joyInfo As JOYINFOEX
    
    joyInfo.dwSize = Len(joyInfo)
    joyInfo.dwFlags = JOY_RETURNALL
    
    errNumber = joyGetPosEx(joyID, joyInfo)
    
    If errNumber = 0 Then
        If (byteVal Or joyInfo.dwButtons) = joyInfo.dwButtons Then
            joyIsButtonStatus = True
        Else
            joyIsButtonStatus = False
        End If
    End If
End Function

Public Function joyIsPOVStatus(joyID As Integer, bytePOV As Long) As Boolean
    'Returns TRUE if joystick POV specified in bytePOV is pressed.
    Dim joyInfo As JOYINFOEX
    
    joyInfo.dwSize = Len(joyInfo)
    joyInfo.dwFlags = JOY_RETURNALL
    
    errNumber = joyGetPosEx(joyID, joyInfo)
    
    If errNumber = 0 Then
        If bytePOV = joyInfo.dwPOV Then
            joyIsPOVStatus = True
        Else
            joyIsPOVStatus = False
        End If
    End If
End Function

Public Sub joyStringtoByte(joyString As String, ByRef joyID As Integer, ByRef bytePOV As Long, ByRef byteVal As Long)
    'Gets the Joystick ID, button byte part and POV byte part off joyString
    
    strArray = Split(joyString, " ")
    joyID = Mid(strArray(0), 4, Len(strArray(0)) - 3)
    bytePOV = -1
    byteVal = -1
    
    Select Case strArray(1)
        Case "HAT": bytePOV = (strArray(2) * 100)
        Case "BTN": btnArray = Split(strArray(2), "+")
                    byteVal = 0
                    For i = 0 To UBound(btnArray)
                        byteVal = byteVal + (2 ^ (CInt(btnArray(i)) - 1))
                    Next
    End Select
End Sub

