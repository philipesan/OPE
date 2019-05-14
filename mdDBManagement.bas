Attribute VB_Name = "mdDBManagement"
Public Sub AbrirDatabase()

Call Log("Abrindo Conexão...")
con.Open strConn
Call Log("Conexão Iniciada com sucesso!")
Exit Sub



End Sub

Public Sub FecharDatabase()



Call Log("Encerrando Conexão...")
con.Close
Call Log("Conexão encerrado com sucesso!")
Exit Sub



End Sub

Public Sub ExecutarQuery(ByVal sSQL As String)

Call Log("Executando query: " & sSQL)
con.Execute sSQL
Call Log("Query Executada com Sucesso!")
Exit Sub




End Sub
