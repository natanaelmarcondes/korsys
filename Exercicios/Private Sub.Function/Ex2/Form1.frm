VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    CentralizarForm Me
    
End Sub

Private Sub CentralizarForm(f As Form)
    
    f.Left = (Screen.Width - f.Width) / 2
    f.Top = (Screen.Height - f.Height) / 2
    
End Sub

