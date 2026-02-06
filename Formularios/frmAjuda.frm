VERSION 5.00
Begin VB.Form frmAjuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sôbre o Sistema"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleMode       =   0  'User
   ScaleWidth      =   8250
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   660
      Left            =   4695
      TabIndex        =   3
      Top             =   3660
      Width           =   1785
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   630
      Left            =   675
      TabIndex        =   2
      Top             =   3510
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   2070
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   450
      Width           =   3285
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   3330
      TabIndex        =   0
      Top             =   2490
      Width           =   1365
   End
End
Attribute VB_Name = "frmAjuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()
                
    'Instanciei na memoria
    Dim frmCad As frmCadUsuarios
    Dim frmCad2 As frmCadUsuarios
    
    Set frmCad = New frmCadUsuarios
    Set frmCad2 = New frmCadUsuarios
    
    frmCad.txtUsername.Text = "ADMIN"
    frmCad.Show
                
    frmCad2.txtUsername.Text = "USUARIO"
    frmCad2.Show
                
    'frmCadUsuarios.Show
        
End Sub
Private Sub Command1_Click()
    
    Dim colUsuarios As Collection
    Set colUsuarios = New Collection

    colUsuarios.Add "João", "j"
    colUsuarios.Add "Maria", "m"
    colUsuarios.Add "Pedro", "p"

    MsgBox colUsuarios("j") ' João




End Sub

Private Sub Command3_Click()
    
    Dim TextBox As clsTextBoxSimples
    Set TextBox = New clsTextBoxSimples
    
    VerificaCampo TextBox
    
End Sub
Private Sub VerificaCampo(Campo As CommandButton)

    If Campo.Text = "" Then
        Campo.Text = "VAZIO"
    End If
    
End Sub

Private Sub Form_Load()
    
'    Dim dicUsuarios As Object
'
'    Set dicUsuarios = CreateObject("Scripting.Dictionary")
'
'    dicUsuarios.Add "Key", "Chave, Fechadura"
'    dicUsuarios.Add "USER", "Usuario Comum"
'
'    If dicUsuarios.Exists("Key") Then
'        MsgBox dicUsuarios("Key")
'    Else
'        MsgBox "Não existe"
'    End If
'
'    Set dicUsuarios = Nothing
    
End Sub

