VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLimpartbm 
      Height          =   345
      Left            =   3240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2040
      Width           =   1410
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar Campo"
      Height          =   795
      Left            =   3075
      TabIndex        =   1
      Top             =   3540
      Width           =   2100
   End
   Begin VB.TextBox txtLimpar 
      Height          =   375
      Left            =   3195
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimpar_Click()
    
    LimparTexto
    
End Sub

Private Sub LimparTexto()
    
    txtLimpar.Text = ""
    txtLimpartbm.Text = ""
    
End Sub
