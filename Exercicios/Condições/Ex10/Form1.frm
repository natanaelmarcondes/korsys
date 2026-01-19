VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Aceito"
      Height          =   1230
      Left            =   1275
      TabIndex        =   2
      Top             =   3495
      Width           =   2940
   End
   Begin VB.CheckBox chkNews 
      Caption         =   "Noticias"
      Height          =   360
      Left            =   1125
      TabIndex        =   1
      Top             =   2295
      Width           =   1560
   End
   Begin VB.CheckBox chkTermos 
      Caption         =   "Termos"
      Height          =   330
      Left            =   1170
      TabIndex        =   0
      Top             =   1680
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
    
    'termos
    If chkTermos.Value = 1 Then
        MsgBox "Você aceitou os termos"
    Else
        MsgBox "Você não aceitou os termos"
    End If
    
    'noticias
    If chkNews.Value = 1 Then
        MsgBox "Você aceitou receber noticias"
    Else
        MsgBox "Você não aceitou receber noticias"
    End If
        
    
End Sub
