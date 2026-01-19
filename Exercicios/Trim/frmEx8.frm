VERSION 5.00
Begin VB.Form frmEx8 
   Caption         =   "Ex8"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCarregar 
      Caption         =   "Carregar"
      Height          =   855
      Left            =   2190
      TabIndex        =   1
      Top             =   3330
      Width           =   1950
   End
   Begin VB.ComboBox cboUsers 
      Height          =   315
      Left            =   2445
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1275
      Width           =   2895
   End
End
Attribute VB_Name = "frmEx8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCarregar_Click()

    Dim i As Integer

    cboUsers.Clear

    For i = 0 To 4
        cboUsers.AddItem strUsers(i)
    Next i


End Sub

Private Sub Form_Load()
    
    strUsers(0) = "Jorge"
    strUsers(1) = "Jorginho"
    strUsers(2) = "Pai do jorge"
    strUsers(3) = "Vô do pai do jorge"
    strUsers(4) = "Ninguem"
    
End Sub
