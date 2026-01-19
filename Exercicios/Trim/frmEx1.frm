VERSION 5.00
Begin VB.Form frmEx1 
   Caption         =   "Ex1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblSystemName 
      Height          =   405
      Left            =   495
      TabIndex        =   0
      Top             =   630
      Width           =   1380
   End
End
Attribute VB_Name = "frmEx1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    lblSystemName.Caption = SYSTEM_NAME

End Sub

