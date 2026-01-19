VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogar 
      Caption         =   "Logar"
      Height          =   915
      Left            =   2835
      TabIndex        =   2
      Top             =   3045
      Width           =   1860
   End
   Begin VB.TextBox txtSenha 
      Height          =   330
      Left            =   3345
      TabIndex        =   1
      Text            =   "Senha"
      Top             =   1125
      Width           =   1200
   End
   Begin VB.TextBox txtLogin 
      Height          =   360
      Left            =   1410
      TabIndex        =   0
      Text            =   "Login"
      Top             =   1110
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogar_Click()
    
    Dim strLogin As String
    Dim strSenha As String
    
    strLogin = "admin"
    strSenha = "123"
    
    If txtLogin.Text = strLogin And txtSenha.Text = strSenha Then
        MsgBox "Acesso liberado", vbInformation
    Else
        MsgBox "Senha ou Login incorreto", vbCritical
    End If
        
    
End Sub
