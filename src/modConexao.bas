Attribute VB_Name = "modConexao"
Option Explicit

' Variável global para reaproveitar a conexão em todo o projeto
Public conn As ADODB.Connection

' Função para abrir a conexão com o SQL Server
Public Function AbrirConexao() As Boolean
    On Error GoTo ErroConexao
    
    ' Se a conexão já estiver aberta, fecha antes de reabrir
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    
    Set conn = New ADODB.Connection
    
    ' String de conexão utilizando o provedor padrão OLE DB do SQL Server
    conn.ConnectionString = "Provider=SQLOLEDB;" & _
                            "Data Source=.\SQLEXPRESS;" & _
                            "Initial Catalog=DB_MiniDespesas;" & _
                            "Integrated Security=SSPI;"
    
    conn.Open
    AbrirConexao = True
    Exit Function

ErroConexao:
    MsgBox "Falha ao conectar ao banco de dados DB_MiniDespesas:" & vbCrLf & _
           Err.Description, vbCritical, "Erro de Conexão"
    AbrirConexao = False
End Function

' Sub para fechar a conexão com segurança ao encerrar o sistema
Public Sub FecharConexao()
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub

