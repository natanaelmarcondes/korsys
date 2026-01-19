VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   2385
      TabIndex        =   2
      Top             =   3855
      Width           =   2775
   End
   Begin VB.TextBox txtNota2 
      Height          =   405
      Left            =   3885
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1470
      Width           =   945
   End
   Begin VB.TextBox txtNota1 
      Height          =   420
      Left            =   1515
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1455
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalcular_Click()
    
    Dim dblNota1 As Double
    Dim dblNota2 As Double
    Dim dblMedia As Double
    
    dblNota1 = Val(txtNota1.Text)
    dblNota2 = Val(txtNota2.Text)
    dblMedia = (dblNota1 + dblNota2) / 2
    
    If dblMedia >= 7 Then
        MsgBox "Passou"
    Else
        MsgBox "Reprovou"
    End If
    
End Sub
