Attribute VB_Name = "mdCriaBanco"
Public Sub CriaBanco(ByVal file As String)
Dim Password As String

Password = "P455c0d3"

Call Log("Conectando ao Provedor de Banco de Dados...")
'Conecta ao Provedor de Banco de dados
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & file & ";Persist Security Info=False"
Call Log("Conectado ao provedor Microsoft.Jet.OLEDB.4.0 com Sucesso!")

'Esse trecho checa se o arquivo de banco já foi criado
Call Log("Verificando existência do banco...")
If Not fso.FileExists(file) Then
    'Criação do Banco de dados
    Call Log("Banco de dados não encontrado, criando...")
    Set dbtemp = CreateDatabase(file, dbLangGeneral)
    Call Log("Banco de dados Criado!")
    Call CriaTabelasPontos
    Call CriaTabelasStatus
    Call CriaTabelasCategorias
    Call CriaTabelasFuncionarios
    Call CriaTabelasCargos
    dbtemp.Close
    Call Log("Populando funcionarios com testes...")
    Call ExportaFuncionarioTeste
End If

End Sub
