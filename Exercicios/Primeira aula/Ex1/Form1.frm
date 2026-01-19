VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Executar"
      Height          =   495
      Left            =   3345
      TabIndex        =   0
      Top             =   4470
      Width           =   1890
   End
   Begin VB.Label lblMensagem 
      Height          =   600
      Left            =   2985
      TabIndex        =   1
      Top             =   1140
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMostrar_Click()

    lblMensagem.Caption = "Evento Click executado"

End Sub

Private Sub Form_Load()

End Sub
