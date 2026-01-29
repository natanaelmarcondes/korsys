Attribute VB_Name = "modGlobal"
Option Explicit

Public Const USR_ADMIN As String = "Admin"
Private Const NUM_DIVISAO As Integer = 2

Public Enum Dificuldade
    Num1 = 1
    Medio = 2
    Hard = 3
    HardCore = 4
End Enum

Public Type typUsuario
    Nome As String
    Email As String
    Senha As String
End Type

Public Type typVisitante
    Nome As String
    Email As String
    Senha As String
    Dificult As Dificuldade
End Type


Public Enum enu_StatusPedido
    PedidoAberto = 1
    PedidoFechado = 2
    PedidoCancelado = 3
End Enum


Public Type typ_Produto
    Codigo As Integer
    Descricao As String
    Preco As Double
End Type

Public Function LerArquivoTxt(strCaminho As String) As String

    Dim intArquivo As Integer
    Dim strLinha As String
    Dim strTexto As String
    
    intArquivo = FreeFile
    
    'Abriu o arquivo todo e injetou dentro da variavel
    Open strCaminho For Input As #intArquivo
    
    strTexto = ""
    
    'Enquanto não for o EOF End Of File (Final do Arquivo) executa e Loop
    Do While Not EOF(intArquivo)
        Line Input #intArquivo, strLinha
        strTexto = strTexto & strLinha & vbCrLf
    Loop
        
    Close #intArquivo
    
    LerArquivoTxt = strTexto
    
End Function
Public Sub GravarArquivoTxt(strCaminho As String, strTexto As String)

    Dim intArquivo As Integer
            
    intArquivo = FreeFile
    
    'Abriu o arquivo todo e injetou dentro da variavel
    Open strCaminho For Output As #intArquivo
    Print #intArquivo, strTexto
            
    Close #intArquivo
            
End Sub
Public Sub HorarioLog(strHora As String)
    
    Dim intHora As Integer
    intHora = FreeFile
    
    Open App.Path & "\log.txt" For Append As #intHora
    Print #intHora, Format(Now, "dd-mm-yyyy hh:nn:ss") & " - " & strHora
    Close #intHora
    
End Sub
Public Sub CenterFormInMDI(frm As Form)

    On Error GoTo TrataErro
    
    Dim mdi As MDIForm
    Dim newLeft As Long
    Dim newTop As Long
    
    'Se for MDI Child, centraliza dentro do MDIForm
    If frm.MDIChild = True Then
        
        Set mdi = frm.Parent
        
        newLeft = (mdi.ScaleWidth - frm.Width) \ NUM_DIVISAO
        newTop = (mdi.ScaleHeight - frm.Height) \ NUM_DIVISAO
        
        If newLeft < 0 Then newLeft = 0
        If newTop < 0 Then newTop = 0
        
        frm.Move newLeft, newTop
    Else
        'Fallback: centraliza na tela
        frm.Left = (Screen.Width - frm.Width) \ NUM_DIVISAO
        frm.Top = (Screen.Height - frm.Height) \ NUM_DIVISAO
    End If
    
    Exit Sub

TrataErro:
    'Fallback extra, caso Parent não esteja acessível
    frm.Left = (Screen.Width - frm.Width) \ NUM_DIVISAO
    frm.Top = (Screen.Height - frm.Height) \ NUM_DIVISAO
End Sub

