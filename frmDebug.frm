VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDebug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Debug Window"
   ClientHeight    =   1812
   ClientLeft      =   2952
   ClientTop       =   636
   ClientWidth     =   5652
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1812
   ScaleWidth      =   5652
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox Debug 
      Height          =   1812
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5652
      _ExtentX        =   9970
      _ExtentY        =   3196
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmDebug.frx":0E42
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
Form1.debugwin.Checked = False
End Sub
