VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "WinScript Program"
   ClientHeight    =   3852
   ClientLeft      =   2004
   ClientTop       =   1896
   ClientWidth     =   4068
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3852
   ScaleWidth      =   4068
   Begin VB.TextBox TxtBox 
      BackColor       =   &H00FFFFFF&
      Height          =   852
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label txt 
      Caption         =   "No"
      Height          =   12
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   612
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Me.Height = Me.Height + 10
End Sub

Private Sub Form_Load()

Me.Height = Me.Height + 10
End Sub

Private Sub Form_Resize()
On Error Resume Next
If txt.Caption = "Yes" Then
TxtBox(1).Width = ScaleWidth
TxtBox(1).Height = ScaleHeight
Else
If txt.Caption = "No" Then

End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

