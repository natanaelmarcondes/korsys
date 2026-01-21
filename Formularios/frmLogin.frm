VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login KorSys"
   ClientHeight    =   7560
   ClientLeft      =   7965
   ClientTop       =   4275
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9225
      TabIndex        =   7
      Top             =   4695
      Width           =   1590
   End
   Begin VB.Timer tmrBloqueio 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5775
      Top             =   6615
   End
   Begin VB.Timer tmrHora 
      Interval        =   1000
      Left            =   795
      Top             =   4605
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   7065
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Senha"
      Top             =   3600
      Width           =   3030
   End
   Begin VB.CommandButton cmdEntrar 
      BackColor       =   &H8000000E&
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7395
      TabIndex        =   2
      Top             =   4680
      Width           =   1590
   End
   Begin VB.TextBox txtLogin 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7080
      TabIndex        =   0
      Text            =   "Username"
      Top             =   2700
      Width           =   2985
   End
   Begin VB.Label lblTextoSenha 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7110
      TabIndex        =   6
      Top             =   3285
      Width           =   2070
   End
   Begin VB.Label lblTextoUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome de usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7095
      TabIndex        =   5
      Top             =   2295
      Width           =   2280
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Data e Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   420
      TabIndex        =   4
      Top             =   5835
      Width           =   2220
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bem vindo ao Login KorSys"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   2745
      TabIndex        =   3
      Top             =   885
      Width           =   6030
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'jorge

Dim intTentativa As Integer
Dim intBlockTimer As Integer
Dim strSenha As String

Private Sub cmdEntrar_Click()
    
    
    'Variaveis de tamanho de texto
    Dim intLenLogin As Integer
    Dim intLenSenha As Integer
    
    intLenLogin = Len(Trim(txtLogin.Text))
    intLenSenha = Len(Trim(txtSenha.Text))
    
    
    'Check de caixa de login vazia
    If Trim(txtLogin.Text) = "" Then
        MsgBox "Informe o usuário!", vbInformation
        txtLogin.SetFocus
        Exit Sub
    End If
    
    
    'Check de caixa de senha vazia
    If Trim(txtSenha.Text) = "" Then
        MsgBox "Informe a senha!", vbInformation
        txtSenha.SetFocus
        Exit Sub
    End If
    
    
    'Check de usuário correto
    If UCase(txtLogin.Text) <> UCase(strLogin) Then
        MsgBox "Usuário inválido!", vbExclamation
        txtLogin.Text = ""
        txtLogin.SetFocus
        intTentativa = intTentativa + 1
    
    'Check de senha correta
    ElseIf txtSenha.Text <> strSenha Then
        MsgBox "Senha inválida!", vbExclamation
        txtSenha.Text = ""
        txtSenha.SetFocus
        intTentativa = intTentativa + 1
    Else
        MDIFormPrincipal.Show 'Abertura do formulario ao validar Usuário e Senha
        Unload Me
        Exit Sub
    End If
    
    
    'Bloquear se as tentativas chegarem a 3
    If intTentativa >= 3 Then
        cmdEntrar.Enabled = False
        intBlockTimer = 10
        MsgBox "Muitas tentativas! Aguarde 10 segundos.", vbExclamation
        tmrBloqueio.Enabled = True
    End If
    
    
    'Checar se tem mais de 3 caracteres no username
    If intLenLogin < 4 Then
        MsgBox "O Username tem de ter mais de 3 caracteres", vbInformation
    ElseIf intLenLogin > 15 Then
        MsgBox "O Username não pode conter mais de 15 caracteres", vbInformation
    End If
    
    
    'Checar se tem mais de 4 caracteres na senha
    If intLenSenha < 5 Then
        MsgBox "A senha tem de ter mais de 4 caracteres", vbInformation
    ElseIf intLenSenha > 8 Then
        MsgBox "A senha não pode ter mais de 8 caracteres", vbInformation
    End If
    
    
End Sub

Private Sub Form_Load()
    
    
    'Carregamento das variaveis
    strLogin = "Admin"
    strSenha = "1234"
    
    
    'Carregamento instantaneo do timer no Formulário
    lblTime.Caption = Format(Now, "dd/mm/yyyy HH:mm:ss")
    lblWelcome.Caption = "Bem vindo ao Login KorSys Ver.:" & App.Major & "." & App.Minor
    
End Sub
Private Sub tmrBloqueio_Timer()
    
    intBlockTimer = intBlockTimer - 1

    If intBlockTimer <= 0 Then
        tmrBloqueio.Enabled = False
        cmdEntrar.Enabled = True
        intTentativa = 0
        MsgBox "Você pode tentar novamente.", vbInformation
    End If
    
End Sub
Private Sub tmrHora_Timer()
    
    
    'Atualização da data em tempo real do formulário
    lblTime.Caption = Format(Now, "dd/mm/yyyy HH:mm:ss")
    
End Sub
Private Sub txtLogin_GotFocus()
    
    txtLogin.Text = ""
    
End Sub
Private Sub txtLogin_LostFocus()
    
    
    'Ao perder o foco colocar a primeira letra em maiuscula
    If Trim(txtLogin.Text) <> "" Then
        'txtLogin.Text = UCase(Left(txtLogin.Text, 1)) & Mid(txtLogin.Text, 2)
        txtLogin.Text = UCase(txtLogin.Text)
    End If

End Sub


Private Sub txtSenha_GotFocus()
    
    txtSenha.Text = ""
    
End Sub
