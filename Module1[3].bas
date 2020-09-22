Attribute VB_Name = "Module1"
Option Explicit

Sub funs()
 Case "msgbox"
            ExecFunction = MsgBox(argList(0), argList(1), argList(2))
        Case "inputbox"
            ExecFunction = InputBox(argList(0), argList(1), argList(2))
        Case "print"
        ScriptForm.Print argList(0)
        Case "print_form"
        ScriptForm.PrintForm
        Case "print_end"
        Printer.EndDoc
        Case "picture"
        ScriptForm.Picture = LoadPicture(argList(0))
        Case "title"
        ScriptForm.Caption = argList(0)
        Case "icon"
        ScriptForm.Icon = LoadPicture(argList(0))
        Case "height"
        ScriptForm.Height = argList(0)
        Case "width"
        ScriptForm.Width = argList(0)
        Case "visible"
        ScriptForm.Visible = argList(0)
        Case "boderstyle"
        ScriptForm.BorderStyle = argList(0)
        Case "enabled"
        ScriptForm.Enabled = argList(0)
End Sub
