VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRelUsuarios 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   10545
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
      Height          =   540
      Left            =   1110
      TabIndex        =   20
      Top             =   4695
      Width           =   1500
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
      Height          =   555
      Left            =   2685
      TabIndex        =   19
      Top             =   4680
      Width           =   1710
   End
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
      Height          =   585
      Left            =   4500
      TabIndex        =   18
      Top             =   4650
      Width           =   1770
   End
   Begin VB.Frame fraUsuario 
      Caption         =   "Usuário"
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
      Height          =   825
      Left            =   255
      TabIndex        =   12
      Top             =   45
      Width           =   10110
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
         Height          =   225
         Left            =   7755
         TabIndex        =   7
         Top             =   465
         Width           =   1095
      End
      Begin VB.ComboBox cboNivel 
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
         Left            =   6585
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   450
         Width           =   1080
      End
      Begin VB.TextBox txtSenha 
         Height          =   300
         Left            =   4665
         TabIndex        =   5
         Top             =   450
         Width           =   1800
      End
      Begin VB.TextBox txtEmail 
         Height          =   300
         Left            =   2760
         TabIndex        =   4
         Top             =   450
         Width           =   1770
      End
      Begin VB.TextBox txtNome 
         Height          =   315
         Left            =   945
         TabIndex        =   3
         ToolTipText     =   "Digite o nome da busca"
         Top             =   450
         Width           =   1635
      End
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
         Height          =   360
         Left            =   9030
         TabIndex        =   8
         Top             =   270
         Width           =   1005
      End
      Begin VB.TextBox txtId 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Digite o nome da busca"
         Top             =   450
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nivel"
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
         Left            =   6645
         TabIndex        =   17
         Top             =   195
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4740
         TabIndex        =   16
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2835
         TabIndex        =   15
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
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
         Left            =   975
         TabIndex        =   14
         Top             =   210
         Width           =   405
      End
      Begin VB.Label lblNome 
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
         Left            =   165
         TabIndex        =   13
         Top             =   210
         Width           =   150
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   6390
      TabIndex        =   10
      Top             =   4620
      Width           =   1890
   End
   Begin VSFlex8LCtl.VSFlexGrid grdUsuarios 
      Height          =   3330
      Left            =   210
      TabIndex        =   9
      Top             =   1125
      Width           =   8040
      _cx             =   14182
      _cy             =   5874
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Frame fraPesquisa 
      Caption         =   "Pesquisa"
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
      Height          =   825
      Left            =   255
      TabIndex        =   11
      Top             =   30
      Width           =   10110
      Begin VB.ComboBox cboCampos 
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2220
      End
      Begin VB.TextBox txtBusca 
         Height          =   315
         Left            =   2550
         TabIndex        =   1
         ToolTipText     =   "Digite o nome da busca"
         Top             =   360
         Width           =   7260
      End
   End
End
Attribute VB_Name = "frmRelUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eUsuarios
    colId = 0
    colNome = 1
    colEmail = 2
    colSenha = 3
    colNivel = 4
    colAtivo = 5
End Enum

Private Sub ListarUsuarios()

    Dim lng_Linha As Long
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    If AbreConexao = False Then
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset
    
    strSQL = ""
    strSQL = "select id,nome,email,senha,nivel,ativo "
    strSQL = strSQL & "from usuarios "
    strSQL = strSQL & "where "
    
    Select Case cboCampos.Text
        
        Case "Id"
            strSQL = strSQL & "id like '%"
        Case "Nome"
            strSQL = strSQL & "Nome like '%"
        Case "Email"
            strSQL = strSQL & "Email like '%"
    End Select
    
    strSQL = strSQL & txtBusca.Text
    strSQL = strSQL & "%'"
   
    
    
    rs.Open strSQL, cn
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "Usuario não encontrado", vbInformation
        Exit Sub
    Else
        
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            grdUsuarios.Rows = grdUsuarios.Rows + 1
            lng_Linha = grdUsuarios.Rows - 1
        
            grdUsuarios.TextMatrix(lng_Linha, 0) = IIf(IsNull(rs!Id), "", rs!Id)
            grdUsuarios.TextMatrix(lng_Linha, 1) = IIf(IsNull(rs!Nome), "", rs!Nome)
            grdUsuarios.TextMatrix(lng_Linha, 2) = IIf(IsNull(rs!Email), "", rs!Email)
            grdUsuarios.TextMatrix(lng_Linha, 3) = IIf(IsNull(rs!Senha), "", rs!Senha)
            grdUsuarios.TextMatrix(lng_Linha, 4) = IIf(IsNull(rs!Nivel), "", rs!Nivel)
            grdUsuarios.TextMatrix(lng_Linha, 5) = IIf(IIf(IsNull(rs!Ativo), "", rs!Ativo), "ATIVO", "INATIVO")
                                    
            grdUsuarios.TextMatrix(lng_Linha, 5) = IIf(rs!Ativo = 1, "ATIVO", "INATIVO")
            
            
            
            If Not IsNull(rs!Ativo) Then
                grdUsuarios.TextMatrix(lng_Linha, 5) = rs!Ativo
                If rs!Ativo = 0 Then
                    grdUsuarios.TextMatrix(lng_Linha, 5) = "INATIVO"
                Else
                    grdUsuarios.TextMatrix(lng_Linha, 5) = "ATIVO"
                End If
                                
            Else
                grdUsuarios.TextMatrix(lng_Linha, 5) = ""
            End If
                        
            rs.MoveNext
        
        Loop
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    FechaConexao
    
End Sub
Private Sub cmdListar_Click()
    
        
End Sub

Private Sub cmdAceitar_Click()
    
    Dim strSQL As String
    
    If AbreConexao = False Then
        Exit Sub
    End If
    
    strSQL = "insert into usuarios (nome, email, senha, nivel, ativo) values ('" & txtNome.Text & "','" & txtEmail.Text & "','" & txtSenha.Text & "','" & cboNivel.Text & "', " & IIf(chkInativo.Value = 1, 0, 1) & ")"
    
    
    cn.Execute strSQL
    
    FechaConexao
    
    MsgBox "Gravação concluida", vbInformation, "Sucesso"
    
    MontaGrid
    CampoPesquisa True
    ListarUsuarios
    
End Sub

Private Sub cmdCancelar_Click()
    
    LimpaCampos
    ComboFill
    CampoPesquisa True
    
End Sub

Private Sub cmdDeletar_Click()
    
    Dim strSQL As String
    
    If AbreConexao = False Then
        Exit Sub
    End If
    
    strSQL = "delete from usuario where id  = " & txtId.Text & ""
    
    cn.Execute strSQL
    
    FechaConexao
    
    MsgBox "Remoção concluida", vbInformation, "Sucesso"
    
    MontaGrid
    CampoPesquisa True
    ListarUsuarios
    
End Sub

Private Sub cmdIncluir_Click()
    
    LimpaCampos
    ComboFill
    CampoPesquisa False
   
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
            
            
    CenterFormInMDI Me, False
    Me.Height = 5880
    Me.Width = 10665
    ComboFill
    
    MontaGrid
    CampoPesquisa True
    ListarUsuarios
    
    
End Sub

Private Sub MontaGrid()
    
    With grdUsuarios
        'Limpa tudo
        .Clear

        .FixedRows = 1
        .FixedCols = 0
    
        .Rows = 1
        .Cols = 6
        
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        
        .FormatString = "Id|Nome|Email|Senha|Nivel|Ativo"
        
        .ColWidth(eUsuarios.colId) = 600
        .ColWidth(eUsuarios.colNome) = 2000
        .ColWidth(eUsuarios.colEmail) = 2000
        .ColWidth(eUsuarios.colSenha) = 1100
        .ColWidth(eUsuarios.colNivel) = 900
        .ColWidth(eUsuarios.colAtivo) = 500
        
    End With
    
        
End Sub

Private Sub ComboFill()
    
    With cboCampos
        .AddItem "Id"
        .AddItem "Nome"
        .AddItem "Email"
        .Text = "Id"
    End With
    
    
    'Nivel
    With cboNivel
        .Clear
        .AddItem "ADMIN"
        .AddItem "USER"
        .Text = "USER"
    End With
    
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub CampoPesquisa(blnAtiva As Boolean)
    
    fraPesquisa.Visible = IIf(blnAtiva, True, False)
    fraUsuario.Visible = IIf(blnAtiva, False, True)
    
End Sub

Private Sub LimpaCampos()
    
    txtId.Text = ""
    txtNome.Text = ""
    txtEmail.Text = ""
    txtSenha.Text = ""
    chkInativo.Value = 0
    
    
End Sub

Private Sub grdUsuarios_DblClick()
    
    Dim WLLng_Row As Long

    ' Linha atual clicada
    WLLng_Row = grdUsuarios.Row
    
    If WLLng_Row < grdUsuarios.FixedRows Then
        Exit Sub
    End If

    txtId.Text = grdUsuarios.TextMatrix(WLLng_Row, eUsuarios.colId)
    txtNome.Text = grdUsuarios.TextMatrix(WLLng_Row, eUsuarios.colNome)
    txtEmail.Text = grdUsuarios.TextMatrix(WLLng_Row, eUsuarios.colEmail)
    txtSenha.Text = grdUsuarios.TextMatrix(WLLng_Row, eUsuarios.colSenha)
    cboNivel.Text = grdUsuarios.TextMatrix(WLLng_Row, eUsuarios.colNivel)
    
    If grdUsuarios.TextMatrix(WLLng_Row, eUsuarios.colAtivo) = "ATIVO" Then
        chkInativo.Value = 0
    Else
        chkInativo.Value = 1
    End If
    
    
End Sub
