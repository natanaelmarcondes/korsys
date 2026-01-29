VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmControles 
   Caption         =   "Anotações"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   9855
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2235
      Left            =   315
      TabIndex        =   7
      Top             =   2970
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   3942
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvwDados 
      Height          =   2160
      Left            =   6825
      TabIndex        =   6
      Top             =   3015
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   3810
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdAceitar 
      Caption         =   "Aceitar"
      Height          =   375
      Left            =   2310
      TabIndex        =   5
      Top             =   1485
      Width           =   1200
   End
   Begin VB.ListBox lstCidades 
      Height          =   2010
      Left            =   4350
      TabIndex        =   4
      Top             =   240
      Width           =   1950
   End
   Begin VB.ComboBox cboDoc 
      Height          =   315
      Left            =   1395
      TabIndex        =   2
      Top             =   300
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   8205
      TabIndex        =   1
      Top             =   480
      Width           =   810
   End
   Begin VB.TextBox txtAnotacao 
      Height          =   390
      Left            =   1065
      TabIndex        =   0
      Top             =   2310
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1395
      TabIndex        =   3
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "frmControles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceitar_Click()
    lstCidades.AddItem cboDoc.Text
End Sub

Private Sub Form_Load()
    
    HorarioLog "Abrindo bloco de anotações"
    
    CenterFormInMDI Me
    
    Call CarregaTipoDoc
    
    CarregaListView
    
    CarregaTreeview
                
End Sub
Private Sub CarregaCidades()

    lstCidades.Clear
    
    lstCidades.AddItem "SÃO PAULO"
    lstCidades.AddItem "RIO DE JANEIRO"
    lstCidades.AddItem "PARAIBA"
    lstCidades.AddItem "PARANÁ"

End Sub
Private Sub CarregaTipoDoc()
    
    'Limpa tudo
    cboDoc.Clear
    
    'Mostra um texto não adiciona
    cboDoc.Text = "DOCUMENTO"
    
    'Aqui adiciona
    cboDoc.AddItem "RG", 0
    cboDoc.AddItem "CPF", 1
    cboDoc.AddItem "CNH", 2
        
End Sub
Private Sub CarregaListView()
    
    With lvwDados
        .View = lvwReport
        .ColumnHeaders.Add , , "Código", 1500
        .ColumnHeaders.Add , , "Nome", 3000

        Dim item As ListItem
        Set item = .ListItems.Add(, , "001")
        item.SubItems(1) = "Produto A"
    End With
End Sub
Private Sub CarregaTreeview()


    Dim nod As Node

    ' Limpa a árvore
    TreeView1.Nodes.Clear

    ' Nó raiz
    Set nod = TreeView1.Nodes.Add(, , "SIS", "Sistema")

    ' Subnível Cadastros
    Set nod = TreeView1.Nodes.Add("SIS", tvwChild, "CAD", "Cadastros")
    TreeView1.Nodes.Add "CAD", tvwChild, "CLI", "Clientes"
        
    TreeView1.Nodes.Add "CAD", tvwChild, "PRO", "Produtos"

    ' Subnível Relatórios
    Set nod = TreeView1.Nodes.Add("SIS", tvwChild, "REL", "Relatórios")
    TreeView1.Nodes.Add "REL", tvwChild, "VEN", "Vendas"

    ' Expande tudo
    TreeView1.Nodes("SIS").Expanded = True

    

End Sub
