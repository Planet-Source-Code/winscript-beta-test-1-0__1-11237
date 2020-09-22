VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Editor"
   ClientHeight    =   1620
   ClientLeft      =   2412
   ClientTop       =   2088
   ClientWidth     =   3168
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3168
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   3132
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   3132
   End
   Begin VB.Label Label2 
      Caption         =   "Picture:"
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Title:"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1212
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form1.Script.SelText = "title(""" + Text1.Text + """)" & Chr$(13) & Chr$(10)
Form1.Script.SelText = "picture(""" + Text2.Text + """)" & Chr$(13) & Chr$(10)
Unload Me
End Sub
