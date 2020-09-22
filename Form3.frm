VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Commands"
   ClientHeight    =   3840
   ClientLeft      =   1320
   ClientTop       =   1248
   ClientWidth     =   4092
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3840
   ScaleWidth      =   4092
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   252
      Left            =   3240
      TabIndex        =   1
      Top             =   3480
      Width           =   852
   End
   Begin VB.TextBox Text1 
      Height          =   3372
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form3.frx":0E42
      Top             =   0
      Width           =   4092
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Text1.SelText = "MsgBox(""TEXT"",""BUTTONVALUE"",""TITLE"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "InputBox(""TEXT"",""TITLE"",""TEXTININPUTBOX"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "PRINT(""TEXTTOPRINT"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "Print_Form" & Chr$(13) & Chr$(10)
    Text1.SelText = "Print_End" & Chr$(13) & Chr$(10)
    Text1.SelText = "Picture(""LOCATIONOFPICTURE"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "Icon(""LOCATIONOFICON"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "Height(""HEIGHT"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "Width(""WIDTH"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "Create TextBox(""LEFT"",""TOP"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "SelText(""TEXTTOADDTOTEXTBOX"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "Unload(""OBJECTOUNLOAD"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "Locked(""TRUEORFALSE"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "NextLine" & Chr$(13) & Chr$(10)
    Text1.SelText = "LCase(""STRINGTOLOWERCASE"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "UCase(""STRINGTOUPPERCASE"")" & Chr$(13) & Chr$(10)
    Text1.SelText = "Show(""TITLE"")"
Text1.SelText = "Hide"
Text1.SelText = "Cls"
Text1.SelText = "Line(""X"",""Y"",-""X"",""Y"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Point (""X"",""Y"")" & Chr$(13) & Chr$(10)
Text1.SelText = "SetForeColor (""RGBHEXCODE"")" & Chr$(13) & Chr$(10)
Text1.SelText = "SetBackColor(""RGBHEXCODE"")" & Chr$(13) & Chr$(10)
Text1.SelText = "GetTitle" & Chr$(13) & Chr$(10)
Text1.SelText = "SetCurrentX(""X"")" & Chr$(13) & Chr$(10)
Text1.SelText = "GetCurrentX" & Chr$(13) & Chr$(10)
Text1.SelText = "SetCurrentY(""Y"")" & Chr$(13) & Chr$(10)
Text1.SelText = "GetCurrentY" & Chr$(13) & Chr$(10)
Text1.SelText = "GetWidth" & Chr$(13) & Chr$(10)
Text1.SelText = "SetWidth(""WIDTH"")" & Chr$(13) & Chr$(10)
Text1.SelText = "GetHeight" & Chr$(13) & Chr$(10)
Text1.SelText = "SetHeight(""HEIGHT"")" & Chr$(13) & Chr$(10)
Text1.SelText = "TextWidth(""WIDTH"")" & Chr$(13) & Chr$(10)
Text1.SelText = "TextHeight(""HEIGHT"")" & Chr$(13) & Chr$(10)
Text1.SelText = "SetFont(""FONTNAME"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Bold" & Chr$(13) & Chr$(10)
Text1.SelText = "Italic" & Chr$(13) & Chr$(10)
Text1.SelText = "Underline" & Chr$(13) & Chr$(10)
Text1.SelText = "Strike" & Chr$(13) & Chr$(10)
Text1.SelText = "Mid(""STRING"",""START"",""LENGTH"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Len(""EXPRESSION"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Chr(""CHARACTERNUMBER"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Asc(""STRING"")" & Chr$(13) & Chr$(10)
Text1.SelText = "ReVerse(""STRINGTOREVERSE"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Cos(""NUMBERASDOUBLE"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Tan(""NUMBERASDOUBLE"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Log(""NUMBERASDOUBLE"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Rnd(""NUMBER"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Randomize(""NUMBER"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Int(""NUMBER"")" & Chr$(13) & Chr$(10)
Text1.SelText = "Printer_Print" & Chr$(13) & Chr$(10)
Text1.SelText = "Finish" & Chr$(13) & Chr$(10)
Text1.SelText = "Shell(""PATHNAME"")" & Chr$(13) & Chr$(10)
End Sub
