VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "WinScript Beta Test 1.0"
   ClientHeight    =   3516
   ClientLeft      =   1584
   ClientTop       =   1908
   ClientWidth     =   6516
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3516
   ScaleWidth      =   6516
   Begin RichTextLib.RichTextBox Script 
      Height          =   3492
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   5292
      _ExtentX        =   9335
      _ExtentY        =   6160
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0E42
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   3516
      Left            =   0
      ScaleHeight     =   3516
      ScaleWidth      =   1092
      TabIndex        =   0
      Top             =   0
      Width           =   1092
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "WinScript Beta Test 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   732
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   972
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   240
         Picture         =   "frmMain.frx":0EF0
         Top             =   120
         Width           =   384
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   240
      Top             =   960
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu jfdhdhkdhsdhskhdskj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As.."
      End
      Begin VB.Menu hgdjhgjhgjg 
         Caption         =   "-"
      End
      Begin VB.Menu MakeEXE 
         Caption         =   "Compile"
         Shortcut        =   {F6}
      End
      Begin VB.Menu kkkkkkk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuundo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuredo 
         Caption         =   "Redo"
         Shortcut        =   {F4}
      End
      Begin VB.Menu yttu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu jfdyhjhdfhdhf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "Project"
      Begin VB.Menu mnurun 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "Compile As..."
         Shortcut        =   ^{F6}
         Visible         =   0   'False
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop"
         Shortcut        =   {F7}
      End
      Begin VB.Menu jkjkdsjksd 
         Caption         =   "-"
      End
      Begin VB.Menu debugwin 
         Caption         =   "Debug"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Beta Tools"
      Begin VB.Menu mnucodedit 
         Caption         =   "Code Editor"
      End
      Begin VB.Menu mnufrmeditor 
         Caption         =   "Form Editor"
      End
      Begin VB.Menu options 
         Caption         =   "Options"
         Visible         =   0   'False
         Begin VB.Menu mnurunopts 
            Caption         =   "Running Options"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu betacmds 
         Caption         =   "Beta Commands"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function File_OpenEXE(File)
Call Shell(File, vbNormalFocus)
End Function
Private Function FileExists(sFilename As String) As Boolean

    If Len(sFilename) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir(sFilename)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Private Sub r_Click()
Dim myLang As New WSL
Form2.Show
myLang.Script = Script.Text
myLang.ScriptExecute
End Sub

Private Sub debug_Click()

End Sub

Private Sub betacmds_Click()
Form3.Show vbModal, Form1
End Sub

Private Sub debugwin_Click()
If debugwin.Checked = False Then
frmDebug.Show
debugwin.Checked = True
Else
If debugwin.Checked = True Then
frmDebug.Hide
debugwin.Checked = False
End If
End If
End Sub

Private Sub Form_Initialize()
On Error Resume Next
frmSplash.Show
Me.Hide

End Sub

Private Sub Form_Resize()
On Error Resume Next
Script.Height = ScaleHeight
Script.Width = ScaleWidth - 1192
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MakeEXE_Click()
On Error GoTo ErrorCatch:
    
    If Not FileExists(App.Path & "\WSC.exe") Then
        MsgBox "Program could not find 'WSC.exe',Unable to continue.", vbCritical, "Error Compiling"
        Exit Sub
    End If
    
    CD.FileName = ""
    CD.Filter = "Executable Files (*.exe)|*.exe|"
    CD.ShowSave
    If CD.FileName <> "" Then
        If FileExists(CD.FileName) Then
            If MsgBox("Overwrite existing file?", vbQuestion + vbYesNo, "WinScript") = vbNo Then
                Exit Sub
            Else
                Kill CD.FileName
            End If
        End If
        
        Dim nFile As Integer
        nFile = FreeFile
        FileCopy App.Path & "\WSC.exe", CD.FileName
        Open CD.FileName For Output As #nFile
        Print #nFile, "|*WSP*|" & Script.Text
        Close #nFile
        MsgBox "File Compiled!", vbInformation, "WinScript"
        
        Dim sTemp As String, sTemp2 As String
        
        Open CD.FileName For Output As #1
        Open App.Path & "\WSC.exe" For Binary As #2
        
        ' Copy data from jelexe into new exe
        While Not EOF(2)
            sTemp = Input$(2000, #2)
            sTemp2 = sTemp2 & sTemp
            Print #1, sTemp2;
            sTemp2 = ""
            If Len(sTemp) > 2000 Then
                sTemp = ""
            End If
        Wend
        
        ' Append the script
        Print #1, "|*WSP*|" & Script.Text
        
        Close #2
        Close #1
        
        
    End If
    
    Exit Sub
ErrorCatch:
    MsgBox "Error has occured: " & Err.Description, vbCritical, "Error"
    Resume Next
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Form1
End Sub

Private Sub mnucodedit_Click()
MsgBox "We're Sorry," & Chr$(13) & Chr$(10) & "But The Code Editor Is Not Availible In The Beta Version On WinScript", vbInformation, "We're Sorry..."
End Sub

Private Sub mnuCompile_Click()
On Error GoTo ErrorCatch:
    
    If Not FileExists(App.Path & "\WSC.exe") Then
        MsgBox "Program could not find 'WSC.exe',Unable to continue.", vbCritical, "Error Compiling"
        Exit Sub
    End If
    
    CD.FileName = ""
    CD.Filter = "Executable Applacation (*.exa)|*.exa|"
    CD.ShowSave
    If CD.FileName <> "" Then
        If FileExists(CD.FileName) Then
            If MsgBox("Overwrite existing file?", vbQuestion + vbYesNo, "WinScript") = vbNo Then
                Exit Sub
            Else
                Kill CD.FileName
            End If
        End If
        
        Dim nFile As Integer
        nFile = FreeFile
        FileCopy App.Path & "\WSC.exe", CD.FileName
        Open CD.FileName For Output As #nFile
        Print #nFile, "|*WSP*|" & Script.Text
        Close #nFile
        MsgBox "File Compiled!", vbInformation, "WinScript"
        
        Dim sTemp As String, sTemp2 As String
        
        Open CD.FileName For Output As #1
        Open App.Path & "\WSC.exe" For Binary As #2
        
        ' Copy data from jelexe into new exe
        While Not EOF(2)
            sTemp = Input$(2000, #2)
            sTemp2 = sTemp2 & sTemp
            Print #1, sTemp2;
            sTemp2 = ""
            If Len(sTemp) > 2000 Then
                sTemp = ""
            End If
        Wend
        
        ' Append the script
        Print #1, "|*WSP*|" & Script.Text
        
        Close #2
        Close #1
        
        
    End If
    
    Exit Sub
ErrorCatch:
    MsgBox "Error has occured: " & Err.Description, vbCritical, "Error"
    Resume Next
End Sub

Private Sub mnuFileNew_Click()
Script.Text = ""
End Sub

Private Sub mnuFileOpen_Click()
    CD.FileName = ""
    CD.Filter = "WinScript Project |*.wsp|"
    CD.ShowOpen
    If CD.FileName <> "" Then
        Dim nFile As Integer
        nFile = FreeFile
        Open CD.FileName For Input As nFile
        Script.Text = Input(LOF(nFile), nFile)
        Close nFile
    End If
End Sub

Private Sub mnuFileSave_Click()
    CD.FileName = ""
    CD.Filter = "WinScript Project|*.wsp|"
    CD.ShowSave
    If CD.FileName <> "" Then
        Dim nFile As Integer
        nFile = FreeFile
        Open CD.FileName For Output As nFile
        Print #nFile, Script.Text
        Close nFile
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    CD.FileName = ""
    CD.Filter = "WinScript Project|*.wsp|All Files|*.*.*|"
    CD.ShowSave
    If CD.FileName <> "" Then
        Dim nFile As Integer
        nFile = FreeFile
        Open CD.FileName For Output As nFile
        Print #nFile, Script.Text
        Close nFile
    End If
End Sub

Private Sub mnufrmeditor_Click()
Form4.Show vbModal, Form1
End Sub

Private Sub mnuquit_Click()
End
End Sub

Private Sub mnuRun_Click()
'If Form5.Option1.Value = False Then
'Dim myLang As New WinS
'Form2.Show
'myLang.Script = Script.Text
'myLang.ScriptExecute
'Else
'If Form5.Option1.Value = True Then
On Error GoTo ErrorCatch:
    
    If Not FileExists(App.Path & "\WSC.exe") Then
        MsgBox "Program could not find 'WSC.exe',Unable to continue.", vbCritical, "Error Compiling"
        Exit Sub
    End If
        Dim nFile As Integer
        nFile = FreeFile
        FileCopy App.Path & "\WSC.exe", "c:\windows\temp\tempprog.exe"
        Open "c:\windows\temp\tempprog.exe" For Output As #nFile
        Print #nFile, "|*WSP*|" & Script.Text
        Close #nFile
        '
        'MsgBox "File Compiled!", vbInformation, "WinScript"
        
        Dim sTemp As String, sTemp2 As String
        
        Open "c:\windows\temp\tempprog.exe" For Output As #1
        Open App.Path & "\WSC.exe" For Binary As #2
        
        ' Copy data from exe into new exe
        While Not EOF(2)
            sTemp = Input$(2000, #2)
            sTemp2 = sTemp2 & sTemp
            Print #1, sTemp2;
            sTemp2 = ""
            If Len(sTemp) > 2000 Then
                sTemp = ""
            End If
        Wend
        
        ' Append the script
        Print #1, "|*WSP*|" & Script.Text
        
        Close #2
        Close #1
        
        
    
Call Shell("c:\windows\temp\tempprog.exe", vbNormalFocus)
Exit Sub
ErrorCatch:
    MsgBox "Error has occured: " & Err.Description, vbCritical, "Error"
    Resume Next


End Sub

Private Sub mnurunopts_Click()
Form5.Show vbModal, Form1
End Sub

Private Sub mnustop_Click()
On Error GoTo Error
Unload Form2
Error:
MsgBox "No Program Running", vbExclamation, "Error"
End Sub

