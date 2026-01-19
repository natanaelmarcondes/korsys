VERSION 5.00
Begin VB.Form frmEx7 
   Caption         =   "Ex7"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExtrair 
      Caption         =   "Extrair"
      Height          =   885
      Left            =   2730
      TabIndex        =   2
      Top             =   3360
      Width           =   2190
   End
   Begin VB.Label lblExtraido 
      Height          =   420
      Left            =   2895
      TabIndex        =   1
      Top             =   2130
      Width           =   1650
   End
   Begin VB.Label lblExtract 
      Caption         =   "Label1"
      Height          =   300
      Left            =   2895
      TabIndex        =   0
      Top             =   1425
      Width           =   1650
   End
End
Attribute VB_Name = "frmEx7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExtrair_Click()
    
    lblExtraido.Caption = Mid(lblExtract.Caption, 5, 5)
    
End Sub

Private Sub Form_Load()
    
    lblExtract.Caption = "PRD-12345-SP"
    
End Sub
