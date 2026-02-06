VERSION 5.00
Begin VB.Form frmAjuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sôbre o Sistema"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleMode       =   0  'User
   ScaleWidth      =   8250
End
Attribute VB_Name = "frmAjuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()
                
    'Instanciei na memoria
    Dim frmCad As frmCadUsuarios
    Dim frmCad2 As frmCadUsuarios
    
    Set frmCad = New frmCadUsuarios
    Set frmCad2 = New frmCadUsuarios
    
    frmCad.txtUsername.Text = "ADMIN"
    frmCad.Show
                
    frmCad2.txtUsername.Text = "USUARIO"
    frmCad2.Show
                
End Sub
Private Sub Command1_Click()
    
    Dim colUsuarios As Collection
    Set colUsuarios = New Collection

    colUsuarios.Add "João", "j"
    colUsuarios.Add "Maria", "m"
    colUsuarios.Add "Pedro", "p"

    MsgBox colUsuarios("j") ' João
End Sub

