VERSION 5.00
Begin VB.Form frmVisualizador 
   Caption         =   "Visualizador de Imagens"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   634
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   939
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   225
      TabIndex        =   4
      Top             =   4410
      Width           =   2115
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   180
      TabIndex        =   3
      Top             =   975
      Width           =   2145
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   195
      TabIndex        =   2
      Top             =   405
      Width           =   2280
   End
   Begin VB.PictureBox picImagem 
      Height          =   7965
      Left            =   2775
      ScaleHeight     =   7905
      ScaleWidth      =   10755
      TabIndex        =   1
      Top             =   345
      Width           =   10815
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   480
      Left            =   11970
      TabIndex        =   0
      Top             =   8895
      Width           =   1950
   End
End
Attribute VB_Name = "frmVisualizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WLBol_SprayAtivo As Boolean
Private Sub Command1_Click()
               
    'Image1.Picture = LoadPicture("C:\Korsys\imagens\exemplo.jpg")
    
End Sub

Private Sub cmdSair_Click()
    
    Unload Me
    
End Sub

Private Sub Dir1_Change()
    
    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
    
    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub File1_Click()
    
    picImagem.Picture = LoadPicture(Dir1.Path & File1.FileName)
    
End Sub

Private Sub Form_Load()
            
    CenterFormInMDI Me, False
    
    Me.Width = 14000
    Me.Height = 10000
    
    File1.Pattern = "*.jpg"
        
End Sub
