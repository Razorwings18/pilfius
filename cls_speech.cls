VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_speech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'API calls for keyboard emulation (active window)
Private Declare Function CharToOem Lib "USER32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Sub keybd_event Lib "USER32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function VkKeyScan Lib "USER32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Long
Private Declare Function MapVirtualKey Lib "USER32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function OemKeyScan Lib "USER32" (ByVal wOemChar As Integer) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'API calls for mouse emulation
Private Declare Sub mouse_event Lib "USER32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'constants for mouse emulation
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Private Const MOUSEEVENTF_WHEEL = &H800 ' mouse wheel
Private Const MOUSEEVENTF_XDOWN = &H100 'X button down
Private Const MOUSEEVENTF_XUP = &H200 'X button up
Private Const WHEEL_DELTA As Long = 120

'API calls for key and mouse capture
Private Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer
'constants for key and mouse capture stuff
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const VK_LBUTTON = &H1 'for Left mouse button
Private Const VK_RBUTTON = &H2 ' for Right mouse button
Private Const VK_MBUTTON = &H4 ' for Middle mouse button
Private Const VK_XBUTTON1 = &H5 ' for mouse button X1
Private Const VK_XBUTTON2 = &H6 ' for mouse button X2

Const m_GrammarId = 10
Const KEYEVENTF_EXTENDEDKEY = &H1

Dim WithEvents RecoContext As SpSharedRecoContext
Attribute RecoContext.VB_VarHelpID = -1
Dim Grammar As ISpeechRecoGrammar
Dim CommandRule As ISpeechGrammarRule
Dim CommandState As ISpeechGrammarRuleState
Dim lastCommandHadInterference As Boolean

Public oCommandSet As cls_command
Private oJoystick As cls_joystick

Public vActivationType As Byte '0=Always active 1=Push-to-activate 2=Push-to-deactivate 3=Push-to-Toggle
Public vActivationKey As Variant 'Holds the keystroke specs when vActivationType is other than 0
Public vNoActions As Boolean 'TRUE if "Don't perform command actions" checkbox selected

'These constants have to be the same as the ones in cls_command
Const KP_COUNT = 6
Const KP_KEY = 0
Const KP_HOLD = 1
Const KP_DELAY = 2
Const KP_REPEAT = 3
Const KP_SHIFT = 4
Const KP_CTRL = 5
Const KP_ALT = 6

Public Event Recognized(Command As String)

Public Sub PauseRecognition()
    RecoContext.Recognizer.State = SRSInactive
End Sub

Public Sub ResumeRecognition()
    RecoContext.Recognizer.State = SRSActive
End Sub

Public Function InitRecognition() As Boolean
    InitRecognition = True
    
    'Initialize the SAPI context
    On Error GoTo EH_INIT
    Set RecoContext = New SpSharedRecoContext
    On Error GoTo 0
    
    'Create single grammar
    Set Grammar = RecoContext.CreateGrammar(m_GrammarId)
    
    'Set CommandRule to be the top (and only) grammar rule, type dynamic
    Set CommandRule = Grammar.Rules.Add("CommandRule", SRATopLevel Or SRADynamic, 2)
    
    'Set rule state for reference
    Set CommandState = CommandRule.AddState
    
    CommandState.AddRuleTransition Nothing, CommandRule, "", 0, 0
    
    CommandState.AddWordTransition Nothing, "", " ", , "", 0, 0
    Exit Function
    
EH_INIT:
    InitRecognition = False
    MsgBox "There is no Speech Recognition Engine installed on your machine. Please, visit the DOWNLOAD section at www.pilfius.com.ar to download one." & vbNewLine & vbNewLine & "See PiLfIuS! documentation for further details. The application will now exit.", vbCritical + vbOKOnly, "NO SPEECH ENGINE INSTALLED"
    mdi_main.TerminateProgram
End Function

Private Sub Class_Initialize()
    Set oJoystick = New cls_joystick
End Sub

Private Sub Class_Terminate()
    Set RecoContext = Nothing
    Set Grammar = Nothing
    Set CommandRule = Nothing
End Sub

Private Sub RecoContext_Interference(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal Interference As SpeechLib.SpeechInterference)
    lastCommandHadInterference = True
End Sub

Private Sub RecoContext_PhraseStart(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
    lastCommandHadInterference = False
End Sub

Private Sub RecoContext_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, _
                                    ByVal RecognitionType As SpeechLib.SpeechRecognitionType, _
                                    ByVal Result As SpeechLib.ISpeechRecoResult)
    Dim currentElement As Integer
    Dim fullCommand As String
    
    fullCommand = ""
    
    If Result.PhraseInfo.Properties.Count > 0 And Not lastCommandHadInterference Then
'''' CARLO's stuff
'Open App.Path & "\output.txt" For Output As #1
'''''''''
        For j = 0 To Result.PhraseInfo.Properties.Count - 1
            If (Result.PhraseInfo.Properties(j).Value > 0) Then
                'ONLY RESPOND IF CONFIDENCE FOR THIS COMMAND IS NOT NEGATIVE WHEN CONFIDENCE THRESHOLD IS ACTIVE
                If (Not oCommandSet.vConfidenceThreshold) Or (oCommandSet.vConfidenceThreshold And Result.PhraseInfo.Properties(j).Confidence >= SECNormalConfidence) Then
'''' CARLO's stuff
'Print #1, Result.PhraseInfo.Properties(j).Name
'''''''''
                    Debug.Print Result.PhraseInfo.Properties(j).Name
                    fullCommand = fullCommand & Result.PhraseInfo.Properties(j).Name & " "
                    'SEND THE COMMAND'S KEYSTROKES
                    If Not vNoActions Then
                        If IsArray(oCommandSet.cKeys(Result.PhraseInfo.Properties(j).Value)) Then
                            For i = 0 To UBound(oCommandSet.cKeys(Result.PhraseInfo.Properties(j).Value), 1)
                                DoSendKey oCommandSet.cKeys(Result.PhraseInfo.Properties(j).Value)(i)
                                Sleep 30
                            Next
                        End If
                    End If
                End If
            End If
        Next
'''' CARLO's stuff
'Close #1
'''''''''
    End If
    
    If Len(fullCommand) > 0 Then RaiseEvent Recognized(fullCommand)
End Sub

Public Sub LoadGrammarItem(vText As String, vValue As Integer)
    'Load vText as new word to recognize, which will return vValue if spoken
    CommandRule.InitialState.AddWordTransition CommandState, vText, " ", , vText, 0, vValue
    
    CommandState.AddWordTransition CommandState, vText, " ", , vText, 0, vValue
End Sub

Public Sub RemoveAllGrammar()
    CommandRule.Clear
    
    'Set rule state for reference
    Set CommandState = CommandRule.AddState
    
    CommandState.AddRuleTransition Nothing, CommandRule, "", 0, 0
    
    CommandState.AddWordTransition Nothing, "", " ", , "", 0, 0
End Sub

Public Sub CommitGrammar()
    'Update grammar
    If CommandRule.InitialState.Transitions.Count > 0 Then
        Grammar.Rules.Commit
    
        'Set CommandRule as the active grammar
        Grammar.CmdSetRuleState "CommandRule", SGDSActive
    End If
End Sub

Private Function mouseButtonRef(mouseBtn As String, mouseAction As Byte) As Integer
    If mouseAction = 0 Then
        Select Case mouseBtn
            Case "MOUSE LEFT": mouseButtonRef = MOUSEEVENTF_LEFTUP
            Case "MOUSE MIDDLE": mouseButtonRef = MOUSEEVENTF_MIDDLEUP
            Case "MOUSE RIGHT": mouseButtonRef = MOUSEEVENTF_RIGHTUP
        End Select
    Else
        Select Case mouseBtn
            Case "MOUSE LEFT": mouseButtonRef = MOUSEEVENTF_LEFTDOWN
            Case "MOUSE MIDDLE": mouseButtonRef = MOUSEEVENTF_MIDDLEDOWN
            Case "MOUSE RIGHT": mouseButtonRef = MOUSEEVENTF_RIGHTDOWN
        End Select
    End If
End Function

Private Function chooseInput(inputType As Variant) As String
    If Mid(inputType, 1, 5) = "MOUSE" Then
        chooseInput = "MOUSE"
    ElseIf Mid(inputType, 1, 3) = "JOY" Then
        chooseInput = "JOY"
    Else
        chooseInput = "KEY"
    End If
End Function

Private Sub DoSendKey(vKeyParams As Variant)
    Dim virtualkey As Long
    Dim scancode As Long
    Dim extended As Boolean
    Dim hadmodifier As Boolean
    Dim inputType As String
    Dim repeatKey As Byte
    
'    Dim oemchar$
'    oemchar$ = " " ' 2 character buffer
    
    inputType = chooseInput(CStr(vKeyParams(KP_KEY)))
    repeatKey = IIf(vKeyParams(KP_REPEAT) > 1, vKeyParams(KP_REPEAT), 1)
    
    If inputType = "KEY" Then
        'IF ACTION IS SEND KEYSTROKE, GET VIRTUAL-KEY AND SCANCODE
        addParam = 0
        scancode = 0
        virtualkey = Keycode2Virtualkey(CStr(vKeyParams(KP_KEY)), extended)
        If scancode = 0 Then scancode = MapVirtualKey(virtualkey, 0)
        If extended Then addParam = KEYEVENTF_EXTENDEDKEY Else addParam = 0
    End If
    
    For loopNr = 1 To repeatKey
        hadmodifier = False
        'press shift
        If vKeyParams(KP_SHIFT) = 1 Then
            keybd_event 16, 42, 0, 0
            hadmodifier = True
        End If
        'press ctrl
        If vKeyParams(KP_CTRL) = 1 Then
            keybd_event 17, 29, 0, 0
            hadmodifier = True
        End If
        'press alt
        If vKeyParams(KP_ALT) = 1 Then
            keybd_event 18, 56, 0, 0
            hadmodifier = True
        End If
        
        If hadmodifier Then Sleep 30
        
        'defusePushedKey MUST be put here after the SLEEP, since some apps reactivate any pushed key
        'in the keyboard after a few milliseconds, even if it's been released by keybd_event.
        If inputType = "KEY" Then defusePushedKey vKeyParams
        
        'press
        Select Case inputType
            Case "KEY": keybd_event virtualkey, scancode, addParam Or 0, 0
            Case "MOUSE": Select Case vKeyParams(KB_KEY)
                            Case "MOUSE WHEELUP": mouse_event MOUSEEVENTF_WHEEL, 0, 0, WHEEL_DELTA, 0
                            Case "MOUSE WHEELDN": mouse_event MOUSEEVENTF_WHEEL, 0, 0, (WHEEL_DELTA * (-1)), 0
                            Case "MOUSE4", "MOUSE5", "MOUSE6", "MOUSE7", "MOUSE8":
                                                mouse_event MOUSEEVENTF_XDOWN, 0, 0, (CLng(Mid(vKeyParams(KB_KEY), 6, 1)) - 3), 0
                            Case Else: mouse_event mouseButtonRef(CStr(vKeyParams(KB_KEY)), 1) Or MOUSEEVENTF_ABSOLUTE, 0, 0, 0, 0
                        End Select
        End Select
        
        'Hold in between press and release
        If vKeyParams(KP_HOLD) > 0 Then
            Sleep vKeyParams(KP_HOLD)
        Else
            Sleep 30
        End If
        
        'release
        Select Case inputType
            Case "KEY": keybd_event virtualkey, scancode, addParam Or 2, 0
            Case "MOUSE":
                        If vKeyParams(KB_KEY) <> "MOUSE WHEELUP" And vKeyParams(KB_KEY) <> "MOUSE WHEELDN" Then
                            Select Case vKeyParams(KB_KEY)
                                Case "MOUSE4", "MOUSE5", "MOUSE6", "MOUSE7", "MOUSE8":
                                    mouse_event MOUSEEVENTF_XUP, 0, 0, (CLng(Mid(vKeyParams(KB_KEY), 6, 1)) - 3), 0
                                Case Else:
                                    mouse_event mouseButtonRef(CStr(vKeyParams(KB_KEY)), 0) Or MOUSEEVENTF_ABSOLUTE, 0, 0, 0, 0
                            End Select
                        End If
        End Select
        
        If hadmodifier Then Sleep 30
        'release shift
        If vKeyParams(KP_SHIFT) = 1 Then
            keybd_event 16, 42, 2, 0
        End If
        'release ctrl
        If vKeyParams(KP_CTRL) = 1 Then
            keybd_event 17, 29, 2, 0
        End If
        'release alt
        If vKeyParams(KP_ALT) = 1 Then
            keybd_event 18, 56, 2, 0
        End If
        
        If inputType = "KEY" Then fusePushedKey
        
        'Delay after key
        If vKeyParams(KP_DELAY) > 0 Then Sleep vKeyParams(KP_DELAY)
        
    Next
    
    allSent = True
End Sub

Public Function Keycode2Virtualkey(KeyCode As String, Optional ByRef extended As Boolean) As Long
    Dim virtualkey As Long
    
    If Len(KeyCode) > 0 Then
        extended = False
        virtualkey = 0
        Select Case KeyCode
            Case "{F1}": virtualkey = 112
            Case "{F2}": virtualkey = 113
            Case "{F3}": virtualkey = 114
            Case "{F4}": virtualkey = 115
            Case "{F5}": virtualkey = 116
            Case "{F6}": virtualkey = 117
            Case "{F7}": virtualkey = 118
            Case "{F8}": virtualkey = 119
            Case "{F9}": virtualkey = 120
            Case "{F10}": virtualkey = 121
            Case "{F11}": virtualkey = 122
            Case "{F12}": virtualkey = 123
            Case "{ENTER}": virtualkey = 13
            Case "{BACKSPACE}": virtualkey = 8
            Case "{ESC}": virtualkey = 27
            Case "{PRNT SCRN}": virtualkey = 44
                                extended = True
            Case "{SCROLL LOCK}": virtualkey = 145
            Case "{PAUSE}": virtualkey = 19
                            extended = True
            Case "{TAB}": virtualkey = 9
            Case "{CAPS LOCK}": virtualkey = 20
            Case "{SPACE}": virtualkey = 32
            Case "{INSERT}": virtualkey = 45
                            extended = True
            Case "{HOME}": virtualkey = 36
                            extended = True
            Case "{PGUP}": virtualkey = 33
                            extended = True
            Case "{PGDN}": virtualkey = 34
                            extended = True
            Case "{DEL}": virtualkey = 46
                            extended = True
            Case "{END}": virtualkey = 35
                            extended = True
            Case "{ARROWUP}": virtualkey = 38
                                extended = True
            Case "{ARROWDOWN}": virtualkey = 40
                                extended = True
            Case "{ARROWLEFT}": virtualkey = 37
                                extended = True
            Case "{ARROWRIGHT}": virtualkey = 39
                                    extended = True
            Case "{NUM LOCK}": virtualkey = 144
                                extended = True
            Case "{KP_0}": virtualkey = 96
            Case "{KP_1}": virtualkey = 97
            Case "{KP_2}": virtualkey = 98
            Case "{KP_3}": virtualkey = 99
            Case "{KP_4}": virtualkey = 100
            Case "{KP_5}": virtualkey = 101
            Case "{KP_6}": virtualkey = 102
            Case "{KP_7}": virtualkey = 103
            Case "{KP_8}": virtualkey = 104
            Case "{KP_9}": virtualkey = 105
            Case "{KP_MULTIPLY}": virtualkey = 106
            Case "{KP_DIVIDE}": virtualkey = 111
                                extended = True
            Case "{KP_SUBTRACT}": virtualkey = 109
            Case "{KP_ADD}": virtualkey = 107
            Case "{KP_DECIMAL}": virtualkey = 110
    '        Case "~": CharToOem Left$("~", 1), oemchar$
    '                    ' Get the scan code for this key
    '                    scancode = OemKeyScan(Asc(oemchar$)) And &HFF
                        
        End Select
        
        If virtualkey = 0 Then virtualkey = VkKeyScan(Asc(KeyCode))
        
        Keycode2Virtualkey = virtualkey
    Else
        Keycode2Virtualkey = 0
    End If
End Function

Public Sub runMicTrainer(parenthWnd As Long)
    If RecoContext.Recognizer.IsUISupported("MicTraining") Then
        RecoContext.Recognizer.DisplayUI parenthWnd, "PiLfIuS!", "MicTraining"
    End If
End Sub

Public Sub runUserTrainer(parenthWnd As Long)
    If RecoContext.Recognizer.IsUISupported("UserTraining") Then
        RecoContext.Recognizer.DisplayUI parenthWnd, "PiLfIuS!", "UserTraining"
    End If
End Sub

Public Sub runProfileProperties(parenthWnd As Long)
    If RecoContext.Recognizer.IsUISupported("RecoProfileProperties") Then
        RecoContext.Recognizer.DisplayUI parenthWnd, "PiLfIuS!", "RecoProfileProperties"
    End If
End Sub

Private Sub defusePushedKey(akeyParams As Variant)
    'RELEASES THE KEY CORRESPONDING TO THE ACTIVATION
    'OTHERWISE, IT WOULD CONFLICT WITH THE KEYSTROKE BEING SENT TO THE PROGRAM
    
    Dim virtualkey As Long
    Dim scancode As Long
    Dim action As Long
    Dim extended As Boolean
    
    If vActivationType = 1 Then
        If Not vActivationKey(KP_KEY) = Empty Then
            If chooseInput(vActivationKey(KP_KEY)) = "KEY" Then
                action = 2
                
                addParam = 0
                virtualkey = Keycode2Virtualkey(CStr(vActivationKey(KP_KEY)), extended)
                scancode = MapVirtualKey(virtualkey, 0)
                If extended Then addParam = KEYEVENTF_EXTENDEDKEY Or action Else addParam = action
                
                'press shift
                If vActivationKey(KP_SHIFT) = 1 And akeyParams(KP_SHIFT) <> 1 Then keybd_event 16, 42, action, 0
                'press ctrl
                If vActivationKey(KP_CTRL) = 1 And akeyParams(KP_CTRL) <> 1 Then keybd_event 17, 29, action, 0
                'press alt
                If vActivationKey(KP_ALT) = 1 And akeyParams(KP_ALT) <> 1 Then keybd_event 18, 38, action, 0
                
                If akeyParams(KP_KEY) <> vActivationKey(KP_KEY) Then keybd_event virtualkey, scancode, addParam, 0
            End If
        End If
    End If
End Sub

Private Sub fusePushedKey()
    'THIS FUNCTION RE-PRESSES THE KEY CORRESPONDING TO THE ACTIVATION
    
    Dim virtualkey As Long
    Dim scancode As Long
    Dim action As Long
    Dim isKeyPad As Boolean
    Dim extended As Boolean
    
    If vActivationType = 1 Then
        If Not vActivationKey(KP_KEY) = Empty Then
            If chooseInput(vActivationKey(KP_KEY)) = "KEY" Then
                If isPressedActivation Then
                    action = 0
                    
                    addParam = 0
                    virtualkey = Keycode2Virtualkey(CStr(vActivationKey(KP_KEY)), extended)
                    scancode = MapVirtualKey(virtualkey, 0)
                    If extended Then addParam = KEYEVENTF_EXTENDEDKEY Or action Else addParam = action
                    
                    keybd_event virtualkey, scancode, addParam, 0
                    
                    'press shift
                    If vActivationKey(KP_SHIFT) = 1 Then keybd_event 16, 42, action, 0
                    'press ctrl
                    If vActivationKey(KP_CTRL) = 1 Then keybd_event 17, 29, action, 0
                    'press alt
                    If vActivationKey(KP_ALT) = 1 Then keybd_event 18, 38, action, 0
                End If
            End If
        End If
    End If
End Sub

Public Function isPressedActivation() As Boolean
    'checks whether the activation input (PTA/PTD/PTT) is pressed
    
    Dim keyCheck As Boolean
    Dim shiftCheck As Boolean
    Dim ctrlCheck As Boolean
    Dim altCheck As Boolean
    
    Dim joyID As Integer
    Dim bytePOV As Long
    Dim byteVal As Long
    
    isPressedActivation = False
    If IsArray(vActivationKey) Then
        If chooseInput(vActivationKey(KP_KEY)) = "JOY" Then
            'check for JOY input
            oJoystick.joyStringtoByte CStr(vActivationKey(KP_KEY)), joyID, bytePOV, byteVal
            
            If (bytePOV > -1) Then
                If oJoystick.joyIsPOVStatus(joyID, bytePOV) Then isPressedActivation = True
            ElseIf (byteVal > -1) Then
                If oJoystick.joyIsButtonStatus(joyID, byteVal) Then isPressedActivation = True
            End If
        ElseIf chooseInput(vActivationKey(KP_KEY)) = "MOUSE" Then
            'check for MOUSE input
            Select Case Mid(vActivationKey(KP_KEY), 6, Len(vActivationKey(KP_KEY)) - 5)
                Case " LEFT": If GetAsyncKeyState(VK_LBUTTON) < 0 Then isPressedActivation = True
                Case " MIDDLE": If GetAsyncKeyState(VK_MBUTTON) < 0 Then isPressedActivation = True
                Case " RIGHT": If GetAsyncKeyState(VK_RBUTTON) < 0 Then isPressedActivation = True
                Case "4": If GetAsyncKeyState(VK_XBUTTON1) < 0 Then isPressedActivation = True
                Case "5": If GetAsyncKeyState(VK_XBUTTON2) < 0 Then isPressedActivation = True
            End Select
        Else
            keyCheck = False
            shiftCheck = True
            ctrlCheck = True
            altCheck = True
            
            'check complete key combo (key+shift+ctrl+alt)
            If Len(vActivationKey(KP_KEY)) > 0 Then
                If GetAsyncKeyState(Keycode2Virtualkey(CStr(vActivationKey(KP_KEY)))) < 0 Then keyCheck = True
            Else
                keyCheck = True
            End If
            If GetAsyncKeyState(vbKeyShift) >= 0 Then
                If vActivationKey(KP_SHIFT) = 1 Then shiftCheck = False
            Else
                If vActivationKey(KP_SHIFT) <> 1 Then shiftCheck = False
            End If
            If GetAsyncKeyState(vbKeyControl) >= 0 Then
                If vActivationKey(KP_CTRL) = 1 Then ctrlCheck = False
            Else
                If vActivationKey(KP_CTRL) <> 1 Then ctrlCheck = False
            End If
            If GetAsyncKeyState(vbKeyMenu) >= 0 Then
                If vActivationKey(KP_ALT) = 1 Then altCheck = False
            Else
                If vActivationKey(KP_ALT) <> 1 Then altCheck = False
            End If
            
            If keyCheck And shiftCheck And ctrlCheck And altCheck Then
                isPressedActivation = True
            End If
        End If
    End If
End Function
