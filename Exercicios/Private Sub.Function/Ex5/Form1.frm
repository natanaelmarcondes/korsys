VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChecar 
      Caption         =   "Ver"
      Height          =   570
      Left            =   2865
      TabIndex        =   1
      Top             =   3045
      Width           =   1440
   End
   Begin VB.TextBox txtNome 
      Height          =   435
      Left            =   2520
      TabIndex        =   0
      Top             =   1305
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CheckTxt() As Boolean
    
    If Trim(txtNome.Text) = "" Then
        CheckTxt = False
    Else
        CheckTxt = True
    End If
    
End Function

Private Sub cmdChecar_Click()
    
    If CheckTxt() Then
        MsgBox "Tem algo"
    Else
        MsgBox "precisa ter algo"
    End If
    
End Sub
