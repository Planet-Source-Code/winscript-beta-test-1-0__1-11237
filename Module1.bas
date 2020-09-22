Attribute VB_Name = "Module1"
Option Explicit
'These are the variables for Undo and Redo
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String

Dim msg
Dim Code1, Code2, Code3
Sub New_Project(Text, vbCrap, Title)
msg = MsgBox(Text, vbYesNo + vbCrap, Title)
If msg = vbYes Then
Form1.text1.Text = ""
Else
Exit Sub
End If
End Sub


Sub QuitMe()
msg = MsgBox("Are You Sure You Want To Quit?", vbYesNo + vbQuestion, "Quit?")
If msg = vbYes Then
End
Else
Exit Sub
End If
End Sub


Sub Run_BodyText()
Dim S
Dim Code1, Code2, Code3
Dim Code
Form2.Show
If S = InStr(Form1.text1.Text, "bodytxt;") Then
Run_SelText
Else
Code1 = InStr(Form1.text1.Text, Chr$(34))
Code2 = InStr(Code1 + 1, Form1.text1.Text, Chr$(34))
Code3 = Mid(Form1.text1.Text, Code1, Code2)
Form1.coder.Caption = Code3
BodyText
Run_SelText
End If
End Sub
Sub run_icon()
Dim S
Dim Code1, Code2, Code3
Dim Code
Form2.Show
If S = InStr(Form1.text1.Text, "icon;") Then
Run_BodyText
Else
Code1 = InStr(Form1.text1.Text, Chr$(34))
Code2 = InStr(Code1 + 1, Form1.text1.Text, Chr$(34))
Code3 = Mid(Form1.text1.Text, Code1, Code2)
Form1.coder.Caption = Code3
Icon
Run_BodyText
End If
End Sub

Sub Run2()
Run_Title
End Sub

Function Run_Title()

Dim S
Dim Code1, Code2, Code3
Dim Code
Form2.Show
If S = InStr(0, Form1.text1.Text, "title;") Then
run_icon
Else
Code1 = InStr(Form1.text1.Text, Chr$(34))
Code2 = InStr(Code1 + 1, Form1.text1.Text, Chr$(34))
Code3 = Mid(Form1.text1.Text, Code1, Code2)
Form1.coder.Caption = Code3
Title
run_icon
End If
End Function
Sub Run_SelText()
Dim S
Dim Code1, Code2, Code3
Dim Code
Form2.Show
If S = InStr(Form1.text1.Text, "seltext;") Then

Else
Code1 = InStr(Form1.text1.Text, Chr$(34))
Code2 = InStr(Code1 + 1, Form1.text1.Text, Chr$(34))
Code3 = Mid(Form1.text1.Text, Code1, Code2)
Form1.coder.Caption = Code3
SelText
End If
End Sub


Sub SelText()
Form2.text1.SelText = Form1.coder.Caption
End Sub

Sub BodyText()
Form2.text1.Text = Form1.coder.Caption
End Sub


Sub Icon()
On Error GoTo ErR1
Form2.Icon = LoadPicture(Form1.coder.Caption)
ErR1:
MsgBox "Interpreter Error," & Chr$(13) & "Error Number: " & Err.Number & Chr$(13) & Err.Description & Chr$(13) & "Ending...", vbInformation
Unload Form2
End Sub



Sub Title()
Form2.Caption = Form1.coder.Caption
End Sub


Sub Save_Project(Text2)
    'Save / Save As
    'basicly the same as open, but your savi
    '     ng
    On Error Resume Next


    Form1.CMDialog1.Filter = Text2


        Form1.CMDialog1.FilterIndex = 1
            'the action is 2, 2 is save(duh----(lol)
            '     )


            Form1.CMDialog1.Action = 2
                Open Form1.CMDialog1.filename For Output As #1
                Print #1, 'put the object that u want saved
                Close #1
            End Sub
 
 
 Sub Open_Project(Text2)
    'Open
    'Note Text2 is the Filter Type(ex:*.*,.t
    '     xt,etc)
    On Error Resume Next


    Form1.CMDialog1.Filter = Text2


        Form1.CMDialog1.FilterIndex = 1
            'the action is one, one = open


            Form1.CMDialog1.Action = 1
                Open Form1.CMDialog1.filename For Input As 1
                'before the '=', but what its going to l
                '     ike if u want a txt file u would do like


                '     Form1.Text1.Text = Input$(LOF(1), 1)
                    Form1.text1.Text = Input$(LOF(1), 1)
                    Close 1
                End Sub


