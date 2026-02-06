Attribute VB_Name = "modBancoDados"
Option Explicit
Public cn As ADODB.Connection
Public Function AbreConexao() As Boolean
            
    On Error GoTo trata_erro
                            
    Screen.MousePointer = vbHourglass   '  Aguarde
    
    AbreConexao = False
            
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then
            AbreConexao = True
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    Set cn = New ADODB.Connection
    
    cn.ConnectionString = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=108.179.193.5;Database=natan291_korsys;User=natan291_root;Password=Korsys@2026;Port=3306;Option=3;"
    
    cn.Open
            
    AbreConexao = True
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
trata_erro:
    Screen.MousePointer = vbDefault
    AbreConexao = False
    MsgBox "Falha ao tentar abrir a conexão do banco de dados: " & Chr(13) & Chr(13) & Err.Number & " " & Err.Description, vbInformation, "Tente novamente"
    
End Function
Public Sub FechaConexao()
        
    If cn.State = adStateOpen Then
        cn.Close
    End If

End Sub
