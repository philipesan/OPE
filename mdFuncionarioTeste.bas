Attribute VB_Name = "mdFuncionarioTeste"
Public Sub ExportaFuncionarioTeste()
Dim nome As String
Dim cargo As Integer

Dim CSQL As String

Call AbrirDatabase

nome = "Ana Patricia Costa Arakaki"
cargo = 1
CSQL = "INSERT INTO funcionarios (nome, cargo)"
CSQL = CSQL & "VALUES('"
CSQL = CSQL & nome & " ','" & cargo
CSQL = CSQL & "')"
Call ExecutarQuery(CSQL)

nome = "Reinaldo Blanco Bareiro"
cargo = 1
CSQL = "INSERT INTO funcionarios (nome, cargo)"
CSQL = CSQL & "VALUES('"
CSQL = CSQL & nome & " ','" & cargo
CSQL = CSQL & "')"
Call ExecutarQuery(CSQL)


nome = "Victor Philipe Arakaki Smirnov"
cargo = 1
CSQL = "INSERT INTO funcionarios (nome, cargo)"
CSQL = CSQL & "VALUES('"
CSQL = CSQL & nome & " ','" & cargo
CSQL = CSQL & "')"
Call ExecutarQuery(CSQL)


nome = "Gerente"
cargo = 1
CSQL = "INSERT INTO cargos (nome)"
CSQL = CSQL & "VALUES('"
CSQL = CSQL & nome
CSQL = CSQL & "')"

Call ExecutarQuery(CSQL)
Call FecharDatabase



End Sub

