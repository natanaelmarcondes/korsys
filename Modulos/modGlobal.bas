Attribute VB_Name = "modGlobal"
Option Explicit

Public strLogin As String

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

