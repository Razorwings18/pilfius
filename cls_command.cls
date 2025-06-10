VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cGroups As New Collection
Public cCommands As New Collection '0=Command Name 1=cGroups position
Private vKey As Variant
Public cKeys As New Collection
Public vConfidenceThreshold As Boolean
Public vActivationType As Byte '0=Always active 1=Push-to-activate 2=Push-to-deactivate 3=Push-to-Toggle
Public vActivationKey As Variant 'Holds the keystroke specs when vActivationType is other than 0

Private HuboKey As Boolean
Private activationOpen As Boolean

Const KP_COUNT = 6
Const KP_KEY = 0
Const KP_HOLD = 1
Const KP_DELAY = 2
Const KP_REPEAT = 3
Const KP_SHIFT = 4
Const KP_CTRL = 5
Const KP_ALT = 6



Public Sub LoadCommands(vFilePath As String)
    Dim groupOpen As String
    Dim cmdOpen As Variant
    Dim KeyParams As Variant 'See KP_ constants for structure
    Dim tinputline As String
    Dim tcomando As String
    Dim posFinComando As Integer
    
    Dim KeyOpen As Boolean
    Dim CommandHasNoKeys As Boolean
    
    KeyOpen = False
    activationOpen = False
    
    vConfidenceThreshold = True
    
    HuboKey = False
    CommandHasNoKeys = False
    groupOpen = ""
    cmdOpen = ""
    
    Open vFilePath For Input As #1
    Do While Not EOF(1)
        'Input #1, tinputline
        tinputline = ""
        newLine = False
        Do While Not EOF(1) And Not newLine
            car = Input(1, #1)
            If car = vbCr Then
                newLine = True
                car = Input(1, #1)
            Else
                tinputline = tinputline & car
            End If
        Loop
        
        If Len(Trim(tinputline)) > 0 Then
            tcomando = ""
            posFinComando = InStr(1, Trim(tinputline), " ")
            If posFinComando > 0 Then
                tcomando = Mid(tinputline, 1, (posFinComando - 1))
            Else
                tcomando = Trim(tinputline)
            End If
            
            If (tcomando = "GROUP" Or tcomando = "COMMAND" Or tcomando = "KEY") Then
                If KeyOpen Then
                    CheckKeyClosure tinputline, KeyParams
                    
                    KeyOpen = False
                ElseIf CommandHasNoKeys And tcomando <> "KEY" Then
                    cKeys.Add 0
                    CommandHasNoKeys = False
                ElseIf activationOpen And tcomando <> "KEY" Then
                    activationOpen = False
                End If
            End If
            
            If tcomando = "ACTIVATION" Then
                activationOpen = True
                
                Select Case Mid(tinputline, posFinComando + 1)
                    Case "AA": vActivationType = 0
                    Case "PTA": vActivationType = 1
                    Case "PTD": vActivationType = 2
                    Case "PTT": vActivationType = 3
                End Select
            End If
            
            If tcomando = "CONFIDENCETHRESHOLD" Then
                If Mid(tinputline, posFinComando + 1) = "1" Then
                    vConfidenceThreshold = True
                Else
                    vConfidenceThreshold = False
                End If
            End If
            
            If tcomando = "GROUP" Then
                groupOpen = Mid(tinputline, 7)
                
                cGroups.Add groupOpen
            End If
            
            If tcomando = "COMMAND" Then
                ReDim cmdOpen(1)
                
                cmdOpen(0) = Mid(tinputline, 9)
                cmdOpen(1) = cGroups.Count
                
                cCommands.Add cmdOpen
                CommandHasNoKeys = True
            End If
            
            If tcomando = "KEY" Then
                CommandHasNoKeys = False
                ReDim KeyParams(KP_COUNT)
                
                If posFinComando > 0 Then
                    KeyParams(KP_KEY) = Mid(tinputline, posFinComando + 1)
                Else
                    KeyParams(KP_KEY) = ""
                End If
                
                KeyOpen = True
            End If
            
            If Trim(Split(tinputline, "=")(0)) = "DELAY" Then
                KeyParams(KP_DELAY) = Trim(Split(tinputline, "=")(1))
            End If
            
            If Trim(Split(tinputline, "=")(0)) = "HOLD" Then
                KeyParams(KP_HOLD) = Trim(Split(tinputline, "=")(1))
            End If
            
            If Trim(Split(tinputline, "=")(0)) = "REPEAT" Then
                KeyParams(KP_REPEAT) = Trim(Split(tinputline, "=")(1))
            End If
            
            If Trim(Split(tinputline, "=")(0)) = "SHIFT" Then
                KeyParams(KP_SHIFT) = Trim(Split(tinputline, "=")(1))
            End If
            
            If Trim(Split(tinputline, "=")(0)) = "CTRL" Then
                KeyParams(KP_CTRL) = Trim(Split(tinputline, "=")(1))
            End If
            
            If Trim(Split(tinputline, "=")(0)) = "ALT" Then
                KeyParams(KP_ALT) = Trim(Split(tinputline, "=")(1))
            End If
        End If
    Loop
    
    If KeyOpen Then
        CheckKeyClosure "", KeyParams
    ElseIf CommandHasNoKeys And tcomando <> "KEY" Then
        cKeys.Add 0
    End If
    Close #1
End Sub

Private Sub CheckKeyClosure(tinputline As String, KeyParams As Variant)
    If activationOpen Then
        'current key belongs to the activation key
        vActivationKey = KeyParams
        
        activationOpen = False
    Else
        'current key belongs to a command
        If HuboKey Then
            ReDim Preserve vKey(UBound(vKey, 1) + 1)
        Else
            ReDim vKey(0)
            HuboKey = True
        End If
        vKey(UBound(vKey, 1)) = KeyParams
        
        If Not Mid(tinputline, 1, 3) = "KEY" Then
            cKeys.Add vKey
            HuboKey = False
        End If
    End If
End Sub

Public Sub addGroup(groupName As String)
    'Adds a new group to the class
    cGroups.Add groupName
End Sub

Public Sub addCommand(commandString As String, groupNumber As Integer)
    Dim cmdOpen As Variant
    
    'Add a new command string
    ReDim cmdOpen(1)
    
    cmdOpen(0) = commandString
    cmdOpen(1) = groupNumber
    
    cCommands.Add cmdOpen
    
    'Add an empty array to cKeys at the corresponding position
    cKeys.Add 0
End Sub

Private Function analyzeKeyName(pkeyString As String) As Variant
    Dim aKeySplit As Variant
    Dim KeyParams As Variant '0=Key 1=Hold 2=Delay 3=Shift 4=Ctrl 5=Alt 6=DestWindow
    
    'Analyze pkeyString and load corresponding data in KeyParams
    ReDim KeyParams(KP_COUNT)
    
    If Mid(pkeyString, 1, 5) <> "MOUSE" And Mid(pkeyString, 1, 3) <> "JOY" Then
        aKeySplit = Split(pkeyString, "+")
        For i = 0 To UBound(aKeySplit)
            If UCase(Trim(aKeySplit(i))) = "SHIFT" Then
                KeyParams(KP_SHIFT) = 1
            ElseIf UCase(Trim(aKeySplit(i))) = "CTRL" Then
                KeyParams(KP_CTRL) = 1
            ElseIf UCase(Trim(aKeySplit(i))) = "ALT" Then
                KeyParams(KP_ALT) = 1
            Else
                KeyParams(KP_KEY) = Trim(aKeySplit(i))
            End If
        Next
    Else
        KeyParams(KP_KEY) = pkeyString
    End If
    
    analyzeKeyName = KeyParams
End Function

Public Sub assignActivationKey(pkeyString As String)
    Dim KeyParams As Variant '0=Key 1=Hold 2=Delay 3=Shift 4=Ctrl 5=Alt 6=DestWindow
    
    KeyParams = analyzeKeyName(pkeyString)
    
    vActivationKey = KeyParams
End Sub

Public Sub addKey(pkeyString As String, pkeyHold As Integer, pkeyDelay As Integer, pCommandNr As Integer, pkeyRepeat As Byte)
    Dim KeyParams As Variant '0=Key 1=Hold 2=Delay 3=Shift 4=Ctrl 5=Alt 6=DestWindow
    Dim tempvKey As Variant
    
    KeyParams = analyzeKeyName(pkeyString)
    KeyParams(KP_HOLD) = pkeyHold
    KeyParams(KP_DELAY) = pkeyDelay
    KeyParams(KP_REPEAT) = pkeyRepeat
    
    'Add new key to the command
    If IsArray(cKeys(pCommandNr)) Then
        tempvKey = cKeys(pCommandNr)
        ReDim Preserve tempvKey(UBound(tempvKey, 1) + 1)
    Else
        ReDim tempvKey(0)
    End If
    
    tempvKey(UBound(tempvKey, 1)) = KeyParams
    
    cKeys.Remove pCommandNr
    If pCommandNr > cKeys.Count Then
        cKeys.Add tempvKey
    Else
        cKeys.Add tempvKey, , pCommandNr
    End If
End Sub

Public Sub eraseGroup(idxGroup)
    Dim tempCmd As Variant
    
    cGroups.Remove idxGroup
    For i = 1 To cCommands.Count
        If cCommands(i)(1) > idxGroup Then
            ReDim tempCmd(1)
            
            tempCmd(0) = cCommands(i)(0)
            tempCmd(1) = cCommands(i)(1) - 1
            
            cCommands.Remove i
            
            If i > cCommands.Count Then
                cCommands.Add tempCmd
            Else
                cCommands.Add tempCmd, , i
            End If
        ElseIf cCommands(i)(1) = idxGroup Then
            ReDim tempCmd(1)
            
            tempCmd(0) = cCommands(i)(0)
            tempCmd(1) = 0
            
            cCommands.Remove i
            
            If i > cCommands.Count Then
                cCommands.Add tempCmd
            Else
                cCommands.Add tempCmd, , i
            End If
        End If
    Next
End Sub

Public Sub eraseCommand(idxCommand)
    cCommands.Remove idxCommand
    cKeys.Remove idxCommand
End Sub


Public Sub eraseKey(idxCommand, idxKey)
    Dim tempKey As Variant
    Dim hadValue As Boolean
    
    hadValue = False
    For i = 0 To UBound(cKeys(idxCommand))
        If i <> idxKey Then
            If hadValue Then
                ReDim Preserve tempKey(UBound(tempKey, 1) + 1)
            Else
                ReDim tempKey(0)
                hadValue = True
            End If
            tempKey(UBound(tempKey)) = cKeys(idxCommand)(i)
        End If
    Next
    cKeys.Remove idxCommand
    If idxCommand > cKeys.Count Then
        cKeys.Add tempKey
    Else
        cKeys.Add tempKey, , idxCommand
    End If
End Sub

Public Sub modifyCommand(idxCommand, pCommand, pGroup)
    ReDim tempCommand(1)
    
    tempCommand(0) = pCommand
    tempCommand(1) = pGroup
    
    cCommands.Remove idxCommand
    
    If idxCommand > cCommands.Count Then
        cCommands.Add tempCommand
    Else
        cCommands.Add tempCommand, , idxCommand
    End If
End Sub
