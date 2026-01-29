VERSION 5.00
Begin VB.Form frmPromocao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promoção"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2000
   ScaleMode       =   0  'User
   ScaleWidth      =   4000
   Begin VB.CommandButton cmdProdutoInfo 
      Caption         =   "Informações do produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1395
      TabIndex        =   7
      Top             =   3660
      Width           =   1335
   End
   Begin VB.CommandButton cmdPedido 
      Caption         =   "Conferir Status do pedido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1380
      TabIndex        =   6
      Top             =   2850
      Width           =   1365
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Calcular Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1395
      TabIndex        =   2
      Top             =   2040
      Width           =   1290
   End
   Begin VB.TextBox txtProduto 
      Height          =   435
      Left            =   1335
      TabIndex        =   0
      Top             =   1230
      Width           =   1155
   End
   Begin VB.Label lblSemPromo 
      AutoSize        =   -1  'True
      Caption         =   "Sem promoção:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2610
      TabIndex        =   5
      Top             =   1320
      Width           =   1710
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5190
      TabIndex        =   4
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label lblPromocao 
      AutoSize        =   -1  'True
      Caption         =   "Promoção : 10% De Desconto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2865
      TabIndex        =   3
      Top             =   465
      Width           =   2895
   End
   Begin VB.Label lblProduto 
      Caption         =   "Produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1350
      TabIndex        =   1
      Top             =   960
      Width           =   675
   End
End
Attribute VB_Name = "frmPromocao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const DESCONTO As Double = 10
Dim StatusPedido As enu_StatusPedido

Private Sub cmdPedido_Click()
    
    StatusPedido = PedidoFechado
    
    MsgBox StatusPedido
    
End Sub

Private Sub cmdProdutoInfo_Click()
    
    Dim typProduto As typ_Produto
    
    typProduto.Codigo = 666
    typProduto.Descricao = "Drogas"
    typProduto.Preco = 29.99
    
    MsgBox "Produto: " & typProduto.Descricao & " Codigo: " & typProduto.Codigo & " Preço: " & typProduto.Preco
    
    
End Sub

Private Sub cmdTotal_Click()
    
    If txtProduto.Text = "" Then
        MsgBox "Tem que ter algo ai"
        Exit Sub
    End If
    
    If Not IsNumeric(txtProduto.Text) Then
        MsgBox "Tem que ser numero"
        Exit Sub
    End If
    
    lblSemPromo = ("Sem Promoção: " & txtProduto.Text)
    lblTotal = ("Total: " & txtProduto.Text - txtProduto.Text * (DESCONTO / 100))
    
End Sub

Private Sub Form_Load()
        
    
    
    CenterFormInMDI Me
    
End Sub
    



