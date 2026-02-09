VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login KorSys"
   ClientHeight    =   7560
   ClientLeft      =   7950
   ClientTop       =   4260
   ClientWidth     =   11415
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmLogin.frx":08CA
   ScaleHeight     =   7560
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExibirSenha 
      Caption         =   "Exibir Senha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10185
      TabIndex        =   7
      Top             =   3630
      Width           =   1095
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   6975
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Senha"
      Top             =   3705
      Width           =   3045
   End
   Begin VB.CommandButton cmdEntrar 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8955
      Picture         =   "frmLogin.frx":544D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4305
      Width           =   1155
   End
   Begin VB.TextBox txtLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6915
      TabIndex        =   0
      Text            =   "nathanjorge@gmail.com"
      Top             =   2805
      Width           =   3120
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este Software foi Desenvolvido por 2N Systems"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   7170
      Width           =   3420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "www.nmarcondes.com.br"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   45
      TabIndex        =   5
      Top             =   7350
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   6540
      TabIndex        =   4
      Top             =   3390
      Width           =   525
   End
   Begin VB.Label lblTextoUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   6540
      TabIndex        =   3
      Top             =   2505
      Width           =   450
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkExibirSenha_Click()
    
    chkExibir
    
End Sub

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
Private Sub cmdSair_Click()
    
    End
    
End Sub

Private Sub Form_Load()
        
    HorarioLog "Iniciado"
            
End Sub
Private Sub txtLogin_LostFocus()
    
    If Trim(txtLogin.Text) <> "" Then
        'txtLogin.Text = UCase(Left(txtLogin.Text, 1)) & Mid(txtLogin.Text, 2)
        txtLogin.Text = LCase(txtLogin.Text)
    End If

End Sub

Private Sub txtSenha_GotFocus()
    
    txtSenha.Text = ""
    
End Sub


Private Sub chkExibir()
    
    If chkExibirSenha.Value = 1 Then
        txtSenha.PasswordChar = ""
    Else
        txtSenha.PasswordChar = "*"
    End If
    
End Sub
