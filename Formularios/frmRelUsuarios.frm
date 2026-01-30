VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmRelUsuarios 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   8025
   Begin VB.CommandButton Command2 
      Caption         =   "Limpar"
      Height          =   360
      Left            =   6570
      TabIndex        =   6
      Top             =   600
      Width           =   1125
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   5265
      TabIndex        =   5
      Top             =   150
      Width           =   1290
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3345
      TabIndex        =   4
      Top             =   150
      Width           =   1710
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   150
      Width           =   1800
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   150
      Width           =   660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adicionar"
      Height          =   360
      Left            =   5340
      TabIndex        =   1
      Top             =   600
      Width           =   1245
   End
   Begin VSFlex8LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   3255
      Left            =   270
      TabIndex        =   0
      Top             =   1110
      Width           =   7230
      _cx             =   12753
      _cy             =   5741
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
    colCodigo = 0
    colNome = 1
    colEmail = 2
    colSenha = 3
End Enum

Dim bolAltera As Boolean
Dim intLinha As Integer
Private Sub VSFlexGrid1_DblClick()
    
    bolAltera = True
    intLinha = VSFlexGrid1.Row
    
    Text1.Text = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, eUsuarios.colCodigo)
    Text2.Text = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, eUsuarios.colNome)
    Text3.Text = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, eUsuarios.colEmail)
    Text4.Text = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, eUsuarios.colSenha)
    
End Sub
Private Sub Command1_Click()
    
    With VSFlexGrid1
        If bolAltera = True Then
            .TextMatrix(intLinha, eUsuarios.colCodigo) = Text1.Text
            .TextMatrix(intLinha, eUsuarios.colNome) = Text2.Text
            .TextMatrix(intLinha, eUsuarios.colEmail) = Text3.Text
            .TextMatrix(intLinha, eUsuarios.colSenha) = Text4.Text
        Else
            .Rows = .Rows + 1 'Cria mais uma linha
            .TextMatrix(.Rows - 1, eUsuarios.colCodigo) = Text1.Text
            .TextMatrix(.Rows - 1, eUsuarios.colNome) = Text2.Text
            .TextMatrix(.Rows - 1, eUsuarios.colEmail) = Text3.Text
            .TextMatrix(.Rows - 1, eUsuarios.colSenha) = Text4.Text
        End If
    End With
    
    bolAltera = False
    intLinha = 0
    
    LimpaTextbox
    
End Sub
Private Sub LimpaTextbox()
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""

End Sub
Private Sub Command2_Click()
    
    MontaGrid
    
End Sub
Private Sub Form_Load()
    
    CenterFormInMDI Me, True
    
    MontaGrid
        
End Sub
Private Sub MontaGrid()
    
    With VSFlexGrid1
        'Limpa tudo
        .Clear

        .FixedRows = 1
        .FixedCols = 0
    
        .Rows = 1
        .Cols = 4
        
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        
        .FormatString = "Código|Nome|Email|Senha"
        
        .ColWidth(eUsuarios.colCodigo) = 1000
        .ColWidth(eUsuarios.colNome) = 2500
        .ColWidth(eUsuarios.colEmail) = 2500
        .ColWidth(eUsuarios.colSenha) = 1000
        
    End With
    
        
End Sub


