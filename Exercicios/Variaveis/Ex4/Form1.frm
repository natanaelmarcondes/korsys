VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Command1"
      Height          =   765
      Left            =   2550
      TabIndex        =   0
      Top             =   3000
      Width           =   2010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExibir_Click()

    Dim strCidade As String
    
    strCidade = "São Paulo"
    
    MsgBox "Você mora na cidade de " & strCidade
    
End Sub
