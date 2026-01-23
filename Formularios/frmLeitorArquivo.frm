VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLeitorArquivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leitor de Arquivo"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9495
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   465
      Left            =   6630
      TabIndex        =   6
      Top             =   6150
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Arquivo de Texto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   150
      TabIndex        =   4
      Top             =   975
      Width           =   9150
      Begin VB.TextBox txtConteudo 
         Height          =   4455
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   330
         Width           =   8745
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Arquivo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   9240
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "Abrir"
         Height          =   420
         Left            =   7830
         TabIndex        =   3
         Top             =   165
         Width           =   1290
      End
      Begin VB.TextBox txtArquivo 
         Height          =   330
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   7545
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   465
      Left            =   7965
      TabIndex        =   0
      Top             =   6150
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdlArquivo 
      Left            =   225
      Top             =   6150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLeitorArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAbrir_Click()
       
    txtArquivo.Text = AbrirArquivo
        
    txtConteudo.Text = LerArquivoTxt(txtArquivo.Text)
    
End Sub
Private Sub cmdLimpar_Click()
    
    AjustaForm
    
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    
    AjustaForm
    
End Sub
Private Sub AjustaForm()
    
    Me.Width = 9510
    Me.Height = 7185
    txtArquivo.Text = ""
    txtConteudo.Text = ""
    
End Sub
Private Function AbrirArquivo() As String
    
    With cdlArquivo
        .DialogTitle = "Abrir Arquivo"
        .Filter = "Arquivos de Texto (*.txt)|*.txt|Todos os Arquivos (*.*)|*.*"
        .ShowOpen
        AbrirArquivo = .FileName
    End With
    
End Function
