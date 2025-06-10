VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm mdi_main 
   BackColor       =   &H0074FFEA&
   Caption         =   "PiLfIuS!"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "mdi_main.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlg_main 
      Left            =   9975
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnu_file_open 
         Caption         =   "Load &Command List..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu_file_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_file_create 
         Caption         =   "&Create new Command List"
         Shortcut        =   ^K
      End
   End
End
Attribute VB_Name = "mdi_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private commandLineParsed As Boolean
Public vVersionNumber As String

Private Sub MDIForm_Load()
    If Not App.PrevInstance Then
        Debug.Print vbNewLine
        commandLineParsed = False
        vVersionNumber = "Ver " & App.Major & "." & App.Minor
        
        frm_main.Show
    Else
        TerminateProgram
    End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = 0 Then
        Me.Width = 12000
        Me.Height = 9000
    ElseIf Me.WindowState = 2 Then
        Me.WindowState = 0
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    TerminateProgram
End Sub

Public Sub mnu_file_create_Click()
    'frm_main.Enabled = False
    Unload frm_main
    fra_main.Show
End Sub

Public Sub mnu_file_open_Click()
    Dim proceedLoad As Boolean
    Dim framainLoaded As Boolean
    
    dlg_main.DialogTitle = "Load PiLfIuS! Command List"
    dlg_main.Filter = ".lcl (PiLfIuS! Command List)|*.lcl"
    dlg_main.FileName = ""
    dlg_main.ShowOpen
    If Len(dlg_main.FileName) > 0 Then
        proceedLoad = True
        
        For Each X In Forms
            If X.Name = "fra_main" Then framainLoaded = True Else framainLoaded = False
        Next
        
        If (framainLoaded) Then
            If MsgBox("This will close your current Command List and discard any changes." & vbNewLine & "Are you sure you want to load a new Command List?", vbYesNo + vbExclamation, "WARNING") = vbNo Then
                proceedLoad = False
            Else
                'If loading is confirmed, unload current Command List
                Unload fra_main
            End If
        End If
        
        If proceedLoad Then
            If Not loadCommandList(dlg_main.FileName) Then
                MsgBox "There has been a problem reading the command list. The file does not seem to exist.", vbCritical + vbOKOnly, "CANNOT LOAD"
            End If
        End If
    End If
End Sub

Public Function loadCommandList(filePath As String) As Boolean
    If Len(Dir(filePath)) > 0 Then
        loadCommandList = True
        
        fra_main.vCommandFile = filePath
        
        'frm_main.Enabled = False
        On Error Resume Next
        Unload frm_main
        fra_main.Show
        
        On Error GoTo 0
    Else
        loadCommandList = False
    End If
End Function

Public Sub commandlineParse()
    Dim posCommand As Integer
    Dim listPath As String
    Dim startCmd As Integer
    Dim endCmd As Integer
    
    If Not commandLineParsed Then
        If Len(Trim(Command$)) > 0 Then
            posCommand = InStr(1, LCase(Command$), "-list:")
            If posCommand > 0 Then
            '-LIST COMMAND IS IN COMMAND-LINE
                startCmd = InStr(posCommand, Command$, """")
                endCmd = InStr(startCmd + 1, Command$, """")
                If startCmd > 0 And endCmd > 0 Then
                    listPath = Mid(Command$, startCmd + 1, ((endCmd - startCmd) - 1))
                    endCmd = InStrRev(listPath, "\")
                    If endCmd > 0 Then
                        If Not loadCommandList(listPath) Then
                            MsgBox "Command list not found!" & vbNewLine & """" & listPath & """" & vbNewLine & vbNewLine & "PiLfIuS! will now exit", vbCritical + vbOKOnly, "COMMAND-LINE ERROR"
                            TerminateProgram
                        End If
                    Else
                        If Not loadCommandList(App.Path & "\" & listPath) Then
                            MsgBox "Command list not found!" & vbNewLine & """" & App.Path & "\" & listPath & """" & vbNewLine & vbNewLine & "PiLfIuS! will now exit", vbCritical + vbOKOnly, "COMMAND-LINE ERROR"
                            TerminateProgram
                        End If
                    End If
                Else
                    MsgBox "-LIST argument error!" & vbNewLine & "Syntax: -LIST:""<commandlist path>""" & vbNewLine & vbNewLine & "PiLfIuS! will now exit", vbCritical + vbOKOnly, "COMMAND-LINE ERROR"
                    TerminateProgram
                End If
            End If
        End If
        commandLineParsed = True
    End If
End Sub

Public Sub TerminateProgram()
    Static unloading As Boolean
    Dim idx As Integer
    
    If unloading Then Exit Sub
    unloading = True
    
    For idx = Forms.Count - 1 To 0 Step -1
        Unload Forms(idx)
    Next idx
    
    unloading = False
End Sub

Public Sub helpWindow(title As String, text As String)
    frm_helpwindow.lbl_title.Caption = title
    frm_helpwindow.lbl_text.Caption = text
    frm_helpwindow.Show 1
End Sub

