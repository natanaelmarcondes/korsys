VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Command1"
      Height          =   1065
      Left            =   1620
      TabIndex        =   2
      Top             =   3615
      Width           =   1770
   End
   Begin VB.TextBox txtNome 
      Height          =   585
      Left            =   690
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   915
      Width           =   1260
   End
   Begin VB.Label lblExibir 
      Caption         =   "Label1"
      Height          =   450
      Left            =   4695
      TabIndex        =   1
      Top             =   1950
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExibir_Click()
    
    Dim strNome As String
    
    strNome = txtNome.Text
    
    lblExibir.Caption = strNome
    
End Sub
