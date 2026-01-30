VERSION 5.00
Begin VB.Form frmAjuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sôbre o Sistema"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleMode       =   0  'User
   ScaleWidth      =   8280
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   525
      Left            =   3060
      TabIndex        =   3
      Top             =   1635
      Width           =   1740
   End
   Begin VB.TextBox Text2 
      Height          =   780
      Left            =   1095
      TabIndex        =   2
      Top             =   2520
      Width           =   5910
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1125
      TabIndex        =   0
      Top             =   585
      Width           =   5820
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1185
      TabIndex        =   4
      Top             =   225
      Width           =   1815
   End
End
Attribute VB_Name = "frmAjuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
            
    Dim intValor As Integer
    
    On Error GoTo trata_erro
          
        
    MsgBox "Vou gerar um erro", vbInformation
    
    intValor = 390000
    
    MsgBox "Beleza , ele já sabe"
    
    Exit Sub
    
trata_erro:

    MsgBox "Ocorreu um erro no sistema:" & Chr(13) & Chr(13) & "Numero do Erro: " & Err.Number & Chr(13) & "Descrição do Erro: " & Err.Description, vbInformation, "Avise o Nathan"
                  
End Sub

Private Sub Command2_Click()
    
    Dim strHost As String
    Dim strPorta As String
    Dim strUsuario As String
    
    strHost = GReg_Ler("KORSYS", "DB", "Host")
    strPorta = GReg_Ler("KORSYS", "DB", "Porta")
    strUsuario = GReg_Ler("KORSYS", "DB", "Usuario")
    
End Sub

Private Sub Form_Load()
        
    CenterFormInMDI Me, True
    
End Sub

Private Sub LimpaCampos()

    On Error Resume Next
    
    
    
End Sub
