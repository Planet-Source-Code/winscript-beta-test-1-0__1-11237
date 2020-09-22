VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Running Options"
   ClientHeight    =   1596
   ClientLeft      =   3468
   ClientTop       =   2688
   ClientWidth     =   2292
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1596
   ScaleWidth      =   2292
   Begin VB.Frame Frame1 
      Caption         =   "Style"
      Height          =   1572
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2292
      Begin VB.OptionButton Option2 
         Caption         =   "In-Client Running(Debug Window)"
         Height          =   492
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   2052
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Compile On Demand(Nothing)"
         Height          =   372
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2052
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   252
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   732
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Me.Hide
End Sub
