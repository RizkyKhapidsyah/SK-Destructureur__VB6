VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form FrmScript 
   Caption         =   "Script Editor"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rt 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   12938
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmScript.frx":0000
   End
End
Attribute VB_Name = "FrmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
If rt.Width - 100 > 0 Then
    rt.Width = Me.Width - 100
    rt.Height = Me.Height - 100
End If
End Sub

