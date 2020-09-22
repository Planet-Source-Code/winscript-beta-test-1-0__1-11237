VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "WinScript "
   ClientHeight    =   2352
   ClientLeft      =   1044
   ClientTop       =   1584
   ClientWidth     =   2940
   ControlBox      =   0   'False
   Height          =   2904
   Icon            =   "Form1.frx":0000
   Left            =   996
   LinkTopic       =   "Form1"
   ScaleHeight     =   2352
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   Top             =   1080
   Width           =   3036
   Begin VB.Label Coder 
      Height          =   504
      Left            =   348
      TabIndex        =   1
      Top             =   324
      Visible         =   0   'False
      Width           =   684
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   696
      Top             =   2304
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3456
      Left            =   12
      TabIndex        =   0
      Top             =   0
      Width           =   3912
      _ExtentX        =   6900
      _ExtentY        =   6096
      _Version        =   393217
      BackColor       =   12632256
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0442
   End
   Begin GradientTitleBar.Gradient Gradient1 
      Left            =   2364
      Top             =   2568
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "Save &As"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuprintersetup 
         Caption         =   "Printe&r Setup"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuundo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuredo 
         Caption         =   "Re&do"
         Shortcut        =   ^Q
      End
      Begin VB.Menu dash11 
         Caption         =   "-"
      End
      Begin VB.Menu mnucut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu dash12 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "Se&lect All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu project 
      Caption         =   "Pro&ject"
      Begin VB.Menu run 
         Caption         =   "&Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu compile 
         Caption         =   "Make E&XA File"
      End
      Begin VB.Menu stop 
         Caption         =   "S&top"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu helpme 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "A&bout"
      End
      Begin VB.Menu help 
         Caption         =   "Help &File"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu min 
      Caption         =   "--"
   End
   Begin VB.Menu bye 
      Caption         =   "X"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'These are the variables for Undo and Redo
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String

Private Sub about_Click()
Form2.Show
End Sub

Private Sub bye_Click()
QuitMe
End Sub

Private Sub Copy_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from Text1 onto the Clipboard
    Clipboard.SetText text1.SelText
    'Sets the Focus to Text1
    text1.SetFocusCopy2
End Sub

Private Sub delete_Click()
text1.SelText = ""
End Sub

Private Sub Form_Load()
Gradient1.GradientForm Me
End Sub


Private Sub Form_Resize()
text1.Height = ScaleHeight
text1.Width = ScaleWidth
End Sub


Private Sub RichTextBox1_Change()

End Sub


Private Sub Option_Click()

End Sub


Private Sub Min_Click()
Me.WindowState = 1
End Sub


Private Sub mnucut_Click()
    'Clears the Clipboard to put text on it
    Clipboard.Clear
    'Sets the Text from Text1 onto the Clipboard
    Clipboard.SetText text1.SelText
    'Deletes the Selected Text on Text1
    text1.SelText = ""
    'Sets the Focus to Text1
    text1.SetFocus
End Sub

Private Sub mnunew_Click()
New_Project "Start A New Project?", vbQuestion, "New Project"
End Sub

Private Sub mnuopen_Click()
Open_Project "WinScript Project (*WSP)|*.WSP|"
End Sub

Private Sub mnuprint_Click()
Form1.Print text1.Text
End Sub

Private Sub mnuprintersetup_Click()
CMDialog1.ShowPrinter
End Sub

Private Sub mnuquit_Click()
QuitMe
End Sub


Private Sub mnuredo_Click()
'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub mnusave_Click()
Save_Project "WinScript Project (*.WSP)|*.WSP|"
End Sub


Private Sub mnusaveas_Click()
Save_Project "All Files (*.*.*)|*.*.*|"
End Sub


Private Sub mnuundo_Click()
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub Paste_Click()
    'Puts the Text from the clipboard into Text1
    text1.SelText = Clipboard.GetText
    'Sets the Focus to Text1
    text1.SetFocus
End Sub

Private Sub Run_Click()
Run2
End Sub

Private Sub selectall_Click()
    'Sets the cursors position to zero
    text1.SelStart = 0
    'Selects the full length of Text1
    text1.SelLength = Len(text1.Text)
    'Sets the Focus to Text1
    text1.SetFocus
End Sub


Private Sub stop_Click()
Unload Form2
End Sub

Private Sub text1_Change()
    'Basically this updates the Undo and Redo
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = Form1.text1.TextRTF
    End If
    
    
    
    Dim S, mMe
    If S = InStr(text1.Text, "title") Then
   mMe = text1.SelStart = text1.SelStart - 5
    text1.SelLength = mMe
    End If
End Sub


