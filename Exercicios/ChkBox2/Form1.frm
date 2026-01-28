VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "Adicionar Produto"
      Height          =   585
      Left            =   525
      TabIndex        =   4
      Top             =   2400
      Width           =   1500
   End
   Begin VB.ListBox lstProdutos 
      Height          =   1815
      Left            =   2430
      TabIndex        =   3
      Top             =   645
      Width           =   2715
   End
   Begin VB.TextBox txtPreco 
      Height          =   330
      Left            =   510
      TabIndex        =   2
      Text            =   "Preço"
      Top             =   1500
      Width           =   1530
   End
   Begin VB.TextBox txtQtd 
      Height          =   315
      Left            =   510
      TabIndex        =   1
      Text            =   "Quantidade"
      Top             =   1080
      Width           =   1515
   End
   Begin VB.TextBox txtProduto 
      Height          =   315
      Left            =   525
      TabIndex        =   0
      Text            =   "Produto"
      Top             =   675
      Width           =   1500
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Left            =   5490
      TabIndex        =   5
      Top             =   1200
      Width           =   570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdicionar_Click()
    
    If Not ValidarCampos() Then Exit Sub
    AdicionarLista
    
End Sub


Private Sub txtPreco_GotFocus()
    
    txtPreco.Text = ""
    
End Sub

Private Sub txtProduto_GotFocus()
    
    txtProduto.Text = ""
    
End Sub

Private Sub txtQtd_GotFocus()
    
    txtQtd.Text = ""
    
End Sub


Private Function ValidarCampos() As Boolean
    
    ValidarCampos = False
    
    If Trim(txtProduto.Text) = "" Then
        MsgBox "É necessario preencher o produto"
        txtProduto.SetFocus
        Exit Function
    End If
    
    If IsNumeric(txtProduto.Text) Then
        MsgBox "Não pode conter numeros"
        txtProduto.SetFocus
        Exit Function
    End If
    
    
    If Trim(txtQtd.Text) = "" Then
        MsgBox "É necessario preencher a quantidade"
        txtQtd.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtQtd.Text) Then
        MsgBox "Quantidade pode ser apenas numeros"
        txtQtd.SetFocus
        Exit Function
    End If
    
    
    If Trim(txtPreco.Text) = "" Then
        MsgBox "É necessario preencher o preço", vbExclamation
        Exit Function
    End If
    
    If Not IsNumeric(txtPreco.Text) Then
        MsgBox "O preço pode ter apenas numeros"
        txtPreco.SetFocus
        Exit Function
    End If
    
    ValidarCampos = True
    
End Function


Public Sub AdicionarLista()
    
    Dim intIndex As Integer
    Dim dblPreco As Double
    Dim intQtd As Integer
    Dim dblTotal As Double

    
    dblPreco = CDbl(txtPreco.Text)
    intQtd = CInt(txtQtd.Text)
    dblTotal = 0
    
    
    lstProdutos.AddItem txtProduto.Text & " Qtd: " & txtQtd.Text & " R$: " & Format(dblPreco, "0.00")
    
    intIndex = lstProdutos.ListCount - 1
    
    lstProdutos.ItemData(intIndex) = CLng(dblPreco * intQtd * 100)
    
    
    For intIndex = 0 To lstProdutos.ListCount - 1
        dblTotal = dblTotal + lstProdutos.ItemData(intIndex) / 100
    Next intIndex
    
    lblTotal.Caption = "Total: R$ " & Format(dblTotal, "0.00")
    
End Sub

