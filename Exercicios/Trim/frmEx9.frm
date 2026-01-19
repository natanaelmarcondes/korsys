VERSION 5.00
Begin VB.Form frmEx9 
   Caption         =   "Ex9"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCArregar 
      Caption         =   "Carregar Nomes"
      Height          =   765
      Left            =   5400
      TabIndex        =   3
      Top             =   3525
      Width           =   1425
   End
   Begin VB.TextBox txtProcurar 
      Height          =   435
      Left            =   2685
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2100
      Width           =   1440
   End
   Begin VB.CommandButton cmddBuscar 
      Caption         =   "Buscar"
      Height          =   840
      Left            =   2460
      TabIndex        =   1
      Top             =   3525
      Width           =   2145
   End
   Begin VB.ComboBox cboEx9 
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1995
   End
End
Attribute VB_Name = "frmEx9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCArregar_Click()
    
    Dim i As Integer
    
    cboEx9.Clear

    For i = 0 To 4
        cboEx9.AddItem strUsuarios(i)
    Next i
    
    
    
End Sub

Private Sub cmddBuscar_Click()
    
    Dim i As Integer
    Dim strNomeDigitado As String
    Dim bolNomeEncontrado As Boolean
    
    
    strNomeDigitado = UCase(Trim(txtProcurar.Text))
    bolNomeEncontrado = False
    
    
    For i = 0 To 4
        If strNomeDigitado = UCase(strUsuarios(i)) Then
            bolNomeEncontrado = True
            Exit For
        End If
    Next i
        If bolNomeEncontrado Then
            MsgBox "Nome encontrado"
        Else
            MsgBox "Nome não encontrado"
    End If
        
    
    
End Sub

Private Sub Form_Load()


    strUsuarios(0) = "Jorge"
    strUsuarios(1) = "Jorginho"
    strUsuarios(2) = "Pai do jorge"
    strUsuarios(3) = "Vô do pai do jorge"
    strUsuarios(4) = "Ninguem"
    
    
    
    
End Sub
