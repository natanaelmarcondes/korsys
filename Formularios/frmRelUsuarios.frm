VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRelUsuarios 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8265
   Begin VB.TextBox txtBusca 
      Height          =   345
      Left            =   2670
      TabIndex        =   3
      ToolTipText     =   "Digite o nome da busca"
      Top             =   135
      Width           =   5190
   End
   Begin VB.ComboBox cboCampos 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   135
      Width           =   2130
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "Listar Usuários"
      Height          =   615
      Left            =   5655
      TabIndex        =   1
      Top             =   4035
      Width           =   1890
   End
   Begin VSFlex8LCtl.VSFlexGrid grdUsuarios 
      Height          =   3330
      Left            =   195
      TabIndex        =   0
      Top             =   615
      Width           =   7620
      _cx             =   13441
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
        
            grdUsuarios.TextMatrix(lng_Linha, 0) = IIf(Isnull(rs!Id), "", rs!Id)
            grdUsuarios.TextMatrix(lng_Linha, 1) = IIf(Isnull(rs!Nome), "", rs!Nome)
            grdUsuarios.TextMatrix(lng_Linha, 2) = IIf(Isnull(rs!Email), "", rs!Email)
            grdUsuarios.TextMatrix(lng_Linha, 3) = IIf(Isnull(rs!Senha), "", rs!Senha)
            grdUsuarios.TextMatrix(lng_Linha, 4) = IIf(Isnull(rs!Nivel), "", rs!Nivel)
            grdUsuarios.TextMatrix(lng_Linha, 5) = IIf(IIf(Isnull(rs!Ativo), "", rs!Ativo), "ATIVO", "INATIVO")
                                    
            grdUsuarios.TextMatrix(lng_Linha, 5) = IIf(rs!Ativo = 1, "ATIVO", "INATIVO")
            
            
            
            If Not Isnull(rs!Ativo) Then
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
    
    MontaGrid
    ListarUsuarios
    
End Sub

Private Sub Form_Load()
            
            
    CenterFormInMDI Me, True
    MontaGrid
    ComboFill
    
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
    
    cboCampos.Clear
    cboCampos.AddItem "Id"
    cboCampos.AddItem "Nome"
    cboCampos.AddItem "Email"
    cboCampos.Text = "Id"
    
End Sub

