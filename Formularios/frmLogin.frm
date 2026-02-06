VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login KorSys"
   ClientHeight    =   7560
   ClientLeft      =   7965
   ClientTop       =   4275
   ClientWidth     =   11415
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":08CA
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
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7095
      TabIndex        =   5
      Top             =   2385
      Width           =   2280
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Data e Hora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   420
      TabIndex        =   4
      Top             =   5835
      Width           =   2220
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bem vindo ao Login KorSys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3180
      TabIndex        =   3
      Top             =   465
      Width           =   5550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bem vindo ao Login KorSys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3150
      TabIndex        =   8
      Top             =   435
      Width           =   5550
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

Private Const NUM_TENTATIVAS As Integer = 3
Private Sub cmdEntrar_Click()
        
    Dim rs As ADODB.Recordset
            
    If AbreConexao = False Then
        Exit Sub
    End If
           
    'usar o que eu preciso
    Set rs = New ADODB.Recordset
    
    rs.Open "select nome,email,senha,nivel,ativo from usuarios where email = '" & Trim(txtLogin.Text) & "'", cn
        
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "Usuario não encontrado", vbInformation
        Exit Sub
    Else
        tUsuarios.Nome = IIf(IsNull(rs!Nome), "VISITANTE", rs!Nome)
        tUsuarios.Email = rs!Email
        
        If IsNull(rs!Senha) Or rs!Senha = "" Then
            MsgBox "Senha em branco, contate o adminstrador", vbInformation
            rs.Close
            Set rs = Nothing
            FechaConexao
            Exit Sub
        Else
            tUsuarios.Senha = rs!Senha
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    FechaConexao
    
    If txtSenha.Text = tUsuarios.Senha Then
        Unload Me
        MDIFormPrincipal.Show
    Else
        MsgBox "Senha inválida", vbInformation
    End If
        
End Sub

Private Sub Form_Load()
        
    'Carregamento das variaveis
    
    Dim Usuario As typUsuario
    Dim Visitante As typVisitante
            
    Usuario.Nome = "Natanael"
    Usuario.Email = "natanael@gmail.com"
    Usuario.Senha = "1234"
        
    Visitante.Nome = "Visitante"
    Visitante.Email = "teste@.comb.r"
    Visitante.Senha = ""
        
    strSenha = "1234"
    
    HorarioLog "Iniciado"
   
    
    'Carregamento instantaneo do timer no Formulário
    lblTime.Caption = Format(Now, "dd/mm/yyyy HH:mm:ss")
    lblWelcome.Caption = "Bem vindo ao Login KorSys Ver.:" & App.Major & "." & App.Minor
    
    Exit Sub
    
trata_erro:

    MsgBox "Ocorreu um erro no sistema:" & Chr(13) & Chr(13) & "Numero do Erro: " & Err.Number & Chr(13) & "Descrição do Erro: " & Err.Description, vbCritical, "Avise o Nathan"
    
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
