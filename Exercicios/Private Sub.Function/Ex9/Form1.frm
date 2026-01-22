VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEntrar 
      Caption         =   "Entrar"
      Height          =   975
      Left            =   2595
      TabIndex        =   2
      Top             =   2805
      Width           =   1500
   End
   Begin VB.TextBox txtSenha 
      Height          =   675
      Left            =   3675
      TabIndex        =   1
      Text            =   "Senha"
      Top             =   1170
      Width           =   1275
   End
   Begin VB.TextBox txtLogin 
      Height          =   570
      Left            =   1605
      TabIndex        =   0
      Text            =   "Login"
      Top             =   1230
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEntrar_Click()
    
    If ValidarConta = False Then
        MsgBox "Login não pode estar vazio e senha deve conter entre 4 e 15 caracteres"
    Else
        MsgBox "Aprovado"
    End If
    
    
End Sub

Private Function ValidarConta() As Boolean
    
    
    If Trim(txtLogin.Text) = "" Then
        ValidarConta = False
        Exit Function
    End If
    
    
    If Len(txtSenha.Text) < 4 Or Len(txtSenha.Text) > 15 Then
        ValidarConta = False
        Exit Function
    End If
    
    
    ValidarConta = True

    
    
End Function
