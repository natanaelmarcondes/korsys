VERSION 5.00
Begin VB.Form frmCadUsuarios 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6210
   ClientLeft      =   10275
   ClientTop       =   5265
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCadUsuarios.frx":0000
   ScaleHeight     =   6210
   ScaleWidth      =   9300
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5385
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Cadastrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3690
      Picture         =   "frmCadUsuarios.frx":1F1B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   7725
      Picture         =   "frmCadUsuarios.frx":27E5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5115
      Width           =   1455
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3705
      TabIndex        =   2
      Top             =   3480
      Width           =   2865
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3735
      TabIndex        =   1
      Top             =   2805
      Width           =   1740
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3735
      TabIndex        =   0
      Top             =   2145
      Width           =   1740
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   6750
      Picture         =   "frmCadUsuarios.frx":30AF
      Top             =   3405
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   5685
      Picture         =   "frmCadUsuarios.frx":3D79
      Top             =   2730
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5670
      Picture         =   "frmCadUsuarios.frx":4A43
      Top             =   2055
      Width           =   480
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de usuários"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   2475
      TabIndex        =   9
      Top             =   675
      Width           =   4290
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1800
      Picture         =   "frmCadUsuarios.frx":570D
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   2730
      TabIndex        =   8
      Top             =   3465
      Width           =   900
   End
   Begin VB.Label lblSSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   2670
      TabIndex        =   7
      Top             =   2775
      Width           =   975
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome de usuário:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   1110
      TabIndex        =   5
      Top             =   2130
      Width           =   2535
   End
End
Attribute VB_Name = "frmCadUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimpar_Click()
    
    LimparCampos
    
End Sub

Private Sub LimparCampos()
    
    txtUsername.Text = ""
    txtSenha.Text = ""
    txtEmail.Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmCadUsuarios = Nothing
    
End Sub
