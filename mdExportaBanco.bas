Attribute VB_Name = "mdExportaBanco"
Public Sub ExportaBancoPonto()
Dim endereco, telefone, semana, abertura, fechamento, operacao, cep As String

Dim CSQL As String

On Error GoTo trata_erro

Call Log("Exportando valores para o Banco de Dados...")

'Tratamento de erro ao tentar abrir o Banco
Call AbrirDatabase

endereco = frmPonto.tbEndereco.text & " " & frmPonto.tbNumero.text
telefone = frmPonto.tbTelefone.text
cep = frmPonto.tbCep.text
abertura = frmPonto.tbAberturaHrs.text & " : " & frmPonto.tbAberturaMins.text
fechamento = frmPonto.tbFechamentoHrs.text & " : " & frmPonto.tbFechamentoMins.text
operacao = abertura & " - " & fechamento
semana = frmPonto.coFimDeSemana.text
gerente = frmPonto.coGerente.text
gerente = Mid(gerente, InStrRev(gerente, " - ") + 3)
CSQL = "INSERT INTO pontos (endereco, telefone, cep,  gerente, hr_operacao, semana)"
CSQL = CSQL & "VALUES('"
CSQL = CSQL & endereco & " ','" & telefone & " ','" & cep & " ','" & gerente & " ','" & operacao & " ','" & semana
CSQL = CSQL & "')"

Call ExecutarQuery(CSQL)
Call FecharDatabase
MsgBox "Ponto cadastrado com sucesso!"
Exit Sub

trata_erro:
Call Log("Erro ao Exportar dados para o Banco, verifique as informações do Formulário")
MsgBox "Erro ao Exportar dados para o Banco, verifique as informações do Formulário", vbCritical
End Sub
Public Sub ExportaBancoStatus()
Dim nome As String

Dim CSQL As String

'On Error GoTo trata_erro

Call Log("Exportando valores para o Banco de Dados...")

'Tratamento de erro ao tentar abrir o Banco
Call AbrirDatabase

nome = frmStatus.tbStatus.text


CSQL = "INSERT INTO status (nome)"
CSQL = CSQL & "VALUES('"
CSQL = CSQL & nome
CSQL = CSQL & "')"

Call ExecutarQuery(CSQL)
Call FecharDatabase
MsgBox "Status cadastrado com sucesso!"
Exit Sub

trata_erro:
Call Log("Erro ao Exportar dados para o Banco, verifique as informações do Formulário")
MsgBox "Erro ao Exportar dados para o Banco, verifique as informações do Formulário", vbCritical
End Sub

Public Sub ExportaBancoCategoria()
Dim nome As String
Dim adicional As Double

Dim CSQL As String

On Error GoTo trata_erro

Call Log("Exportando valores para o Banco de Dados...")

'Tratamento de erro ao tentar abrir o Banco
Call AbrirDatabase

nome = frmCategoria.tbNome.text
adicional = frmCategoria.tbAdicional.text

CSQL = "INSERT INTO categorias (nome, adicional)"
CSQL = CSQL & "VALUES('"
CSQL = CSQL & nome & " ','" & adicional
CSQL = CSQL & "')"

Call ExecutarQuery(CSQL)
Call FecharDatabase
MsgBox "Categoria cadastrada com sucesso!"

Exit Sub

trata_erro:
Call Log("Erro ao Exportar dados para o Banco, verifique as informações do Formulário")
MsgBox "Erro ao Exportar dados para o Banco, verifique as informações do Formulário", vbCritical

End Sub
Public Sub ExportaBancoCargo()
Dim nome As String
Dim salario As Double
Dim vAdmin, vRh As Integer

Dim CSQL As String

On Error GoTo trata_erro

Call Log("Exportando valores para o Banco de Dados...")

'Tratamento de erro ao tentar abrir o Banco
Call AbrirDatabase

nome = frmCargo.tbNome.text
salario = frmCargo.tbSalario.text
vAdmin = frmCargo.ckAdmin.Value
vRh = frmCargo.ckRh.Value

CSQL = "INSERT INTO cargos (nome, salario, acesso_admin, acesso_rh)"
CSQL = CSQL & "VALUES('"
CSQL = CSQL & nome & " ','" & salario & " ','" & vAdmin & " ','" & vRh
CSQL = CSQL & "')"

Call ExecutarQuery(CSQL)
Call FecharDatabase
MsgBox "Cargo cadastrado com sucesso!"

Exit Sub

trata_erro:
Call Log("Erro ao Exportar dados para o Banco, verifique as informações do Formulário")
MsgBox "Erro ao Exportar dados para o Banco, verifique as informações do Formulário", vbCritical

End Sub
