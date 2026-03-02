VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRegSetores 
   Caption         =   "Registro de Setores"
   ClientHeight    =   5565
   ClientLeft      =   9780
   ClientTop       =   4590
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   8235
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4800
      TabIndex        =   8
      Top             =   4845
      Width           =   1400
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   6630
      TabIndex        =   9
      Top             =   4845
      Width           =   1400
   End
   Begin VB.CommandButton cmdDeletar 
      Caption         =   "Deletar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3285
      TabIndex        =   7
      Top             =   4845
      Width           =   1400
   End
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "Atualizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1770
      TabIndex        =   6
      Top             =   4845
      Width           =   1400
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   270
      TabIndex        =   5
      Top             =   4845
      Width           =   1400
   End
   Begin VSFlex8LCtl.VSFlexGrid grdSetores 
      Height          =   3420
      Left            =   240
      TabIndex        =   13
      Top             =   1185
      Width           =   7725
      _cx             =   13626
      _cy             =   6032
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fraIncluir 
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   840
      Left            =   240
      TabIndex        =   14
      Top             =   135
      Width           =   7770
      Begin VB.CommandButton cmdAceitar 
         Caption         =   "Aceitar"
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
         Left            =   6735
         TabIndex        =   4
         Top             =   255
         Width           =   915
      End
      Begin VB.CheckBox chkInativo 
         Caption         =   "Inativo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5850
         TabIndex        =   3
         Top             =   465
         Width           =   810
      End
      Begin VB.TextBox txtSetorDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   450
         Width           =   2760
      End
      Begin VB.TextBox txtSetNome 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   450
         Width           =   1920
      End
      Begin VB.TextBox txtID 
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
         Height          =   285
         Left            =   165
         TabIndex        =   0
         Top             =   450
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição do Setor"
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
         Left            =   3030
         TabIndex        =   17
         Top             =   210
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Setor"
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
         Left            =   990
         TabIndex        =   16
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Id"
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
         Left            =   180
         TabIndex        =   15
         Top             =   210
         Width           =   150
      End
   End
   Begin VB.Frame fraBusca 
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   840
      Left            =   240
      TabIndex        =   18
      Top             =   150
      Width           =   7770
      Begin VB.ComboBox cboOpcoes 
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
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   330
         Width           =   1380
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
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
         Left            =   6735
         TabIndex        =   12
         Top             =   255
         Width           =   915
      End
      Begin VB.TextBox txtBuscar 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1755
         TabIndex        =   11
         Top             =   300
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmRegSetores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eSetores
    
    colId = 0
    colNome = 1
    colDescricao = 2
    colAtivo = 3
End Enum

Dim bolAltera As Boolean

Private Sub cmdAceitar_Click()


    
    If bolAltera = True Then
        cn.Execute strAltera
    Else
        cn.Execute strNovo
    End If
    
End Sub

Private Sub cmdAtualizar_Click()
    
    Me.Caption = "Cadastro de Setores" & " - Alterando"
    LimpaCampos
    ComboFill
    CampoPesquisa False
    bolAltera = True
    
End Sub

Private Sub cmdIncluir_Click()
    
    Me.Caption = "Cadastro de Setores" & " - Incluindo"
    LimpaCampos
    ComboFill
    CampoPesquisa False
    bolAltera = False
    
    
End Sub

Private Sub cmdSair_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    CenterFormInMDI Me, False
    Me.Height = 6030
    Me.Width = 8355
    
    MontaGrid
    CampoPesquisa True
    ListarSetores
    ComboFill
    
End Sub

Private Sub CampoPesquisa(blnAtiva As Boolean)
    
    fraBusca.Visible = IIf(blnAtiva, True, False)
    fraIncluir.Visible = IIf(blnAtiva, False, True)
    
End Sub

Private Sub ComboFill()
    
    With cboOpcoes
        .AddItem "Id"
        .AddItem "Nome"
        .Text = "Id"
    End With
 
End Sub

Private Sub MontaGrid()
    
    With grdSetores
    
        'Limpa tudo
        .Clear
        
        'cria os fixos
        .FixedRows = 1
        .FixedCols = 0
        
        'cria quantas colunas vai ter
        .Rows = 1
        .Cols = 4
        
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        
        'formato da linha
        .FormatString = "Id|Nome|Descrição|Status"
        
        'tamanho das colunas
        .ColWidth(eSetores.colId) = 600
        .ColWidth(eSetores.colNome) = 2000
        .ColWidth(eSetores.colDescricao) = 3000
        .ColWidth(eSetores.colAtivo) = 500
        
    End With
    
        
End Sub


Private Sub ListarSetores()

    Dim lng_Linha As Long
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    If AbreConexao = False Then
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "select id,nome,descricao,ativo "
    strSQL = strSQL & "from setores "
    strSQL = strSQL & "where "
    
    Select Case cboOpcoes.Text
        
        Case "Id"
            strSQL = strSQL & "id like '%"
        Case "Nome"
            strSQL = strSQL & "Nome like '%"
    End Select
    
    strSQL = strSQL & txtBuscar.Text
    strSQL = strSQL & "%'"
   
    
    
    rs.Open strSQL, cn
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "Setor não encontrado", vbInformation
        Exit Sub
    Else
        
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            grdSetores.Rows = grdSetores.Rows + 1
            lng_Linha = grdSetores.Rows - 1
        
            grdSetores.TextMatrix(lng_Linha, 0) = IIf(IsNull(rs!id), "", rs!id)
            grdSetores.TextMatrix(lng_Linha, 1) = IIf(IsNull(rs!Nome), "", rs!Nome)
            grdSetores.TextMatrix(lng_Linha, 2) = IIf(IsNull(rs!Descricao), "", rs!Descricao)
            'tratamento null para ativo/inativo
            If IsNull(rs!Ativo) Then
                grdSetores.TextMatrix(lng_Linha, 3) = ""
                ElseIf rs!Ativo = 1 Then
                    grdSetores.TextMatrix(lng_Linha, 3) = "ATIVO"
                Else
                    grdSetores.TextMatrix(lng_Linha, 3) = "INATIVO"
            End If

                    
            rs.MoveNext
        
        Loop
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    FechaConexao
    
End Sub
