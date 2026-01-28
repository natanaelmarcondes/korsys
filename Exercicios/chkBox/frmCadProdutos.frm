VERSION 5.00
Begin VB.Form frmCadProdutos 
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5895
      TabIndex        =   13
      Top             =   3915
      Width           =   1650
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8205
      TabIndex        =   11
      Top             =   6210
      Width           =   1410
   End
   Begin VB.CommandButton cmdLista 
      Caption         =   "Adicionar na Lista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3255
      TabIndex        =   10
      Top             =   1815
      Width           =   1635
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   870
      TabIndex        =   9
      Top             =   3885
      Width           =   1575
   End
   Begin VB.TextBox txtPreco 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   720
      TabIndex        =   8
      Text            =   "Preço"
      Top             =   1815
      Width           =   1785
   End
   Begin VB.ListBox lstProdutos 
      Height          =   2010
      Left            =   6015
      TabIndex        =   7
      Top             =   645
      Width           =   3705
   End
   Begin VB.TextBox txtQuantidade 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   720
      TabIndex        =   2
      Text            =   "Quantidade"
      Top             =   1215
      Width           =   1785
   End
   Begin VB.TextBox txtDescrição 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5985
      TabIndex        =   5
      Text            =   "Descrição"
      Top             =   2895
      Width           =   1620
   End
   Begin VB.CheckBox chkAceitarTermos 
      Caption         =   "Aceito os termos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3045
      TabIndex        =   4
      Top             =   1215
      Width           =   2100
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   840
      TabIndex        =   0
      Top             =   3270
      Width           =   1590
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "Adicionar"
      Enabled         =   0   'False
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
      Left            =   825
      TabIndex        =   6
      Top             =   2610
      Width           =   1605
   End
   Begin VB.ComboBox cboItens 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2895
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   675
      Width           =   2865
   End
   Begin VB.TextBox txtAdicionarProduto 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   720
      TabIndex        =   1
      Text            =   "Adicionar Produto"
      Top             =   645
      Width           =   1785
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
      Left            =   6360
      TabIndex        =   12
      Top             =   3510
      Width           =   570
   End
End
Attribute VB_Name = "frmCadProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AceitarTermos()
    
    If chkAceitarTermos.Value = 0 Then
        cmdAdicionar.Enabled = False
    Else
        cmdAdicionar.Enabled = True
    End If
    
End Sub


Private Sub chkAceitarTermos_Click()
    
    If chkAceitarTermos.Value = 0 Then
        cmdAdicionar.Enabled = False
    Else
        cmdAdicionar.Enabled = True
    End If
    
End Sub

Private Sub cmdAdicionar_Click()
    
    AdicionarItem
    
End Sub
Private Sub Limpeza()
    
    chkAceitarTermos.Value = 0
    txtAdicionarProduto.Text = ""
    txtQuantidade.Text = ""
    txtPreco.Text = ""
    txtAdicionarProduto.SetFocus
    
End Sub
Private Sub AdicionarItem()
    
    'Validação de estar vazio
    If Trim(txtAdicionarProduto.Text) = "" Then
        MsgBox "Tem que ter algo", vbExclamation
        txtAdicionarProduto.SetFocus
        Exit Sub
    End If
    
    'validar vazio na quantidade
    If Trim(txtQuantidade.Text) = "" Then
        MsgBox "Também tem q ter algo.", vbExclamation
        txtQuantidade.SetFocus
        Exit Sub
    End If
     
    'validar preço não ta vazio
    If Trim(txtPreco.Text) = "" Then
        MsgBox "aqui precisa ter o preço", vbExclamation
        Exit Sub
    End If
    
    
    'adiciona item na cbobox
    cboItens.AddItem txtAdicionarProduto.Text & " | " & txtQuantidade.Text & " - " & txtPreco.Text & "R$"
    
End Sub


Private Sub cmdLimpar_Click()
    
    Limpeza
    
End Sub

Private Sub cmdLista_Click()
    
    'checa se a cbobox n ta vazia se estiver n adiciona
    If Trim(cboItens.Text) = "" Then
        MsgBox "Necessario conter algo", vbExclamation
        Exit Sub
    End If
    
    lstProdutos.AddItem cboItens.Text
    
End Sub

Private Sub cmdRemover_Click()
    
    RemoverItem
    
End Sub

Private Sub cmdSair_Click()
    
    If MsgBox("Deseja sair?", vbQuestion + vbYesNo) = vbYes Then End
    
End Sub

Private Sub cmdTotal_Click()
    
    TotalPrecos
    
End Sub

Private Sub Form_Load()
    
    cboItens.Clear
    
End Sub


Private Sub txtAdicionarProduto_GotFocus()
    
    txtAdicionarProduto.Text = ""
    
End Sub


Private Sub txtQuantidade_GotFocus()
    
    txtQuantidade.Text = ""
    
End Sub
    
    
Private Sub txtAdicionarProduto_KeyPress(KeyAscii As Integer)

    'Tira numeros comecei a fazer, esse aqui veio do chat, n tenho ideia do numero das teclas
    If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or _
            (KeyAscii >= 97 And KeyAscii <= 122) Or _
            KeyAscii = 32 Or _
            KeyAscii = 8) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)

    'Tira letras
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or _
        KeyAscii = 8) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
    'Só numero
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    
    'Ou virgulal ou ponto
    If (KeyAscii = 44 Or KeyAscii = 46) Then
        If InStr(txtPreco.Text, ",") = 0 And InStr(txtPreco.Text, ".") = 0 Then
            Exit Sub
        End If
    End If

    '8 = apagar
    If KeyAscii = 8 Then Exit Sub

    'trava o resto
    KeyAscii = 0
    
End Sub


Private Sub txtAdicionarProduto_Change()
    
    'não pode digitar se não colocar algo no produto
    If Trim(txtAdicionarProduto.Text) <> "" Then
        txtQuantidade.Enabled = True
    Else
        txtQuantidade.Enabled = False
        txtQuantidade.Text = ""
    End If
    
End Sub

Private Sub txtQuantidade_Change()
    
    'não pode digitar se não colocar na quantidade também
    If Trim(txtQuantidade.Text) <> "" Then
        txtPreco.Enabled = True
    Else
        txtPreco.Enabled = False
        txtPreco.Text = ""
    End If
    
End Sub

Private Sub RemoverItem()
    
    'check de seleção
    If cboItens.ListIndex = -1 Then
        MsgBox "Selecione um item para removeer", vbExclamation
        Exit Sub
    End If
    
    
    cboItens.RemoveItem cboItens.ListIndex
    
    
    cboItens.Text = ""
    
End Sub

Private Sub TotalPrecos()
    
    Dim i As Integer
    Dim dblTotal As Double
    Dim strPartes() As String
    
    dblTotal = 0
    
    For i = 0 To lstProdutos.ListCount - 1
        strPartes = Split(lstProdutos.List(i), "Sub:")
         dblTotal = dblTotal + CDbl(Trim(strPartes(1)))
    Next i
    
    
    
End Sub








