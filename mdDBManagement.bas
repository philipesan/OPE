Attribute VB_Name = "mdDBManagement"
Public Sub AbrirDatabase()

On Error GoTo trata_erro

'Tratamento de erro ao tentar abrir o Banco
Call Log("Abrindo Conex�o...")
con.Open strConn
Call Log("Conex�o Iniciada com sucesso!")
Exit Sub

trata_erro:
Call Log("Erro ao Conectar ao tentar Abrir o Banco de Dados!, verificar Conex�o!")
MsgBox "Erro ao Conectar ao tentar Abrir o Banco de Dados! Verifique o arquivo de Log", vbCritical


End Sub

Public Sub FecharDatabase()

On Error GoTo tratamento_erro

'Tratamento de erro ao tentar abrir o Banco
Call Log("Encerrando Conex�o...")
con.Close
Call Log("Conex�o encerrado com sucesso!")
Exit Sub

tratamento_erro:
Call Log("Erro ao Conectar ao tentar Encerrar o Banco de Dados!, verificar Conex�o!")
MsgBox "Erro ao Conectar ao tentar Fechar o Banco de Dados! Verifique o arquivo de Log", vbCritical


End Sub

Public Sub ExecutarQuery(ByVal sSQL As String)

On Error GoTo trata_erro

'Tratamento de erro ao tentar abrir o Banco
Call Log("Executando query: " & sSQL)
con.Execute sSQL
Call Log("Query Executada com Sucesso!")
Exit Sub

trata_erro:
Call Log("Erro ao Conectar ao tentar Executar Query!, verificar Conex�o!")
MsgBox "Erro ao Conectar ao tentar Executar Query! Verifique o arquivo de Log", vbCritical


End Sub
