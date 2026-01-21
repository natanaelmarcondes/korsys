VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "msCOMctl.OCX"
Begin VB.MDIForm MDIFormPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8685
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13755
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8310
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   1164
      ButtonWidth     =   1376
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Usuarios"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "Cadastro"
      Begin VB.Menu mnuCadUsuarios 
         Caption         =   "Usuários"
      End
      Begin VB.Menu mnuCadClientes 
         Caption         =   "Clientes"
      End
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "Sôbre"
      Begin VB.Menu mnuTela 
         Caption         =   "Tela"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "MDIFormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuCadUsuarios_Click()
    frmCadUsuarios.Show
End Sub

Private Sub mnuSair_Click()
    End
End Sub

Private Sub mnuTela_Click()
    frmAjuda.Show
End Sub
