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
Public Function ExecutarRetornoID(ByVal sSQL As String) As Integer
    Dim retorno As Integer
    Call Log("Executando query: " & sSQL)
    con.Execute sSQL, , adCmdText + adExecuteNoRecords
    Set rs = con.Execute("SELECT @@Identity", , adCmdText)
    Call Log("Query Executada com Sucesso! ID retornado: " & Val(rs.Fields(0).value))
    retorno = Val(rs.Fields(0).value)
    ExecutarRetornoID = retorno
End Function
