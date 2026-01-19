VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Command1"
      Height          =   1020
      Left            =   3705
      TabIndex        =   0
      Top             =   3015
      Width           =   2100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExibir_Click()

    Dim strNome As String
    
    strNome = "Nathan"
    
    MsgBox strNome

End Sub
