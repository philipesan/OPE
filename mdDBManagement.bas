Attribute VB_Name = "mdDBManagement"
Public Sub AbrirDatabase()

Call Log("Abrindo Conex�o...")
con.Open strConn
Call Log("Conex�o Iniciada com sucesso!")
Exit Sub



End Sub

Public Sub FecharDatabase()



Call Log("Encerrando Conex�o...")
con.Close
Call Log("Conex�o encerrado com sucesso!")
Exit Sub



End Sub

Public Sub ExecutarQuery(ByVal sSQL As String)

Call Log("Executando query: " & sSQL)
con.Execute sSQL
Call Log("Query Executada com Sucesso!")
Exit Sub




End Sub
