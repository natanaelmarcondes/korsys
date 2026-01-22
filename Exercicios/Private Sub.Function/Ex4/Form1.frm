VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ativar"
      Height          =   630
      Left            =   4365
      TabIndex        =   2
      Top             =   3810
      Width           =   1290
   End
   Begin VB.TextBox txtTexto 
      Height          =   390
      Left            =   2895
      TabIndex        =   1
      Top             =   1185
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Desativar"
      Height          =   780
      Left            =   5865
      TabIndex        =   0
      Top             =   3720
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    HabilitarCampo txtTexto, False
    
End Sub

Private Sub Command2_Click()
    
    HabilitarCampo txtTexto, True
    
End Sub

Private Sub HabilitarCampo(t As TextBox, b As Boolean)
    
    t.Enabled = b
    
End Sub


