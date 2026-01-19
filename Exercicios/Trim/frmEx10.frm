VERSION 5.00
Begin VB.Form frmEx10 
   Caption         =   "Ex10"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Carrega nome"
      Height          =   645
      Left            =   6135
      TabIndex        =   3
      Top             =   1995
      Width           =   1140
   End
   Begin VB.CommandButton cmdEx10 
      Caption         =   "Checagem"
      Height          =   720
      Left            =   2925
      TabIndex        =   2
      Top             =   3540
      Width           =   2220
   End
   Begin VB.TextBox txtEx10 
      Height          =   525
      Left            =   2685
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1965
      Width           =   1890
   End
   Begin VB.ComboBox cboEx10 
      Height          =   315
      Left            =   2085
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   930
      Width           =   1995
   End
End
Attribute VB_Name = "frmEx10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strUsers(4) As String

Private Sub cmdEx10_Click()
    
    Dim strTexto As String
    Dim strResult As String
    
    strTexto = Trim(txtEx10.Text)
    
    Select Case cboEx10.Text
        
        Case "Maiusculo"
            strResult = UCase(strTexto)
        Case "Minusculo"
            strResult = LCase(strTexto)
        
        Case "Tamanho do Texto"
            strResult = Len(strTexto)
        
        Case "Três primeira letras"
            If Len(strTexto) > 3 Then
                strResult = Left(strTexto, 3)
            Else
                strResult = strTexto
            End If
        
        Case "Três ultimas letrass"
            If Len(strTexto) <= 3 Then
                strResult = Right(strTexto, 3)
            Else
                strResult = strTexto
            End If
        
        Case Else
            strResult = "Tem nada ai n"
    
    End Select
    
        MsgBox strResult
    
End Sub

Private Sub cmdLoad_Click()
    
    Dim i As Integer

    cboEx10.Clear

    For i = 0 To 4
        cboEx10.AddItem strUsers(i)
    Next i


    
End Sub

Private Sub Form_Load()
    
    strUsers(0) = "Jorge"
    strUsers(1) = "Jorginho"
    strUsers(2) = "Pai do jorge"
    strUsers(3) = "Vô do pai do jorge"
    strUsers(4) = "Ninguem"
    
    
    
End Sub
