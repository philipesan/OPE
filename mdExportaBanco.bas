Attribute VB_Name = "mdExportaBanco"
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
    
    CSQL = "INSERT INTO categorias (nome, adicional, flag)"
    CSQL = CSQL & "VALUES('"
    CSQL = CSQL & nome & " ','" & adicional
    CSQL = CSQL & "', 0)"
    
    Call ExecutarQuery(CSQL)
    Call FecharDatabase
    MsgBox "Categoria cadastrada com sucesso!"
    
    Exit Sub
    
trata_erro:
    Call Log("Err.Source & " - " & Err.Description")
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
    vAdmin = frmCargo.ckAdmin.value
    vRh = frmCargo.ckRh.value
    
    CSQL = "INSERT INTO cargos (nome, salario, acesso_admin, acesso_rh, flag)"
    CSQL = CSQL & "VALUES('"
    CSQL = CSQL & nome & " ','" & salario & " ','" & vAdmin & " ','" & vRh
    CSQL = CSQL & "', 0)"
    
    Call ExecutarQuery(CSQL)
    Call FecharDatabase
    MsgBox "Cargo cadastrado com sucesso!"
    
    Exit Sub
    
trata_erro:
    Call Log("Err.Source & " - " & Err.Description")
    MsgBox "Erro ao Exportar dados para o Banco, verifique as informações do Formulário", vbCritical

End Sub
Public Sub ExportaBancoServico()
    Dim nome, descricao As String
    Dim valor As Double
    
    Dim CSQL As String
    
    'On Error GoTo trata_erro
    
    Call Log("Exportando valores para o Banco de Dados...")
    
    'Tratamento de erro ao tentar abrir o Banco
    Call AbrirDatabase
    
    nome = frmServico.tbNome.text
    descricao = frmServico.tbDescricao.text
    valor = frmServico.tbValor.text
    
    CSQL = "INSERT INTO servicos (nome, descricao, preco, flag)"
    CSQL = CSQL & "VALUES('"
    CSQL = CSQL & nome & " ','" & descricao & " ','" & valor
    CSQL = CSQL & "', 0)"
    
    Call ExecutarQuery(CSQL)
    Call FecharDatabase
    MsgBox "Serviço cadastrado com sucesso!"
    
    Exit Sub
    
trata_erro:
    Call Log(Err.Source & " - " & Err.Description)
    MsgBox "Erro ao Exportar dados para o Banco, verifique as informações do Formulário", vbCritical

End Sub
Public Sub ExportaBancoFuncionario()
Dim nome As String, senha As String
Dim cargo As Integer
    Dim CSQL As String
    
    senha = frmFuncionario.tbSenha.text
    
    'Converte a senha para Hex MD5'
    Call Log("Escrevendo senha em Hash MD5")
    senha = UCase(MD5.DigestStrToHexStr(senha))
    MsgBox senha
    
    
    On Error GoTo trata_erro
    
    Call Log("Exportando valores para o Banco de Dados...")
    
    'Tratamento de erro ao tentar abrir o Banco
    Call AbrirDatabase
    
    nome = frmFuncionario.tbNome.text
    cargo = frmFuncionario.coCargo.ListIndex + 1
    
    CSQL = "INSERT INTO funcionarios (nome, cargo, senha, flag)"
    CSQL = CSQL & "VALUES('"
    CSQL = CSQL & nome & " ','" & cargo & " ','" & senha
    CSQL = CSQL & "', 0)"
    
    Call ExecutarQuery(CSQL)
    Call FecharDatabase
    MsgBox "Funcionario cadastrado com sucesso!"
    
    Exit Sub
    
trata_erro:
    Call Log("Err.Source & " - " & Err.Description")
    MsgBox "Erro ao Exportar dados para o Banco, verifique as informações do Formulário", vbCritical

End Sub
Public Function ExportaBancoOrdem() As Integer

    Dim Usuario, Cliente, Placa, Marca, Modelo, Telefone, Email, Documento As String
    Dim Hora As Date
    Dim Categoria, Status, resposta As Integer
    Dim Valor_Total As Double

    Dim CSQL As String
    
    Hora = Now
    Usuario = Val(Left(frmTelaCadastros.lbSessao.Caption, InStr(frmTelaCadastros.lbSessao.Caption, "-") - 1))
    Cliente = frmCriarVenda.tbNome.text
    Placa = frmCriarVenda.tbPlaca.text
    Marca = frmCriarVenda.tbMarca.text
    Modelo = frmCriarVenda.tbModelo.text
    Telefone = frmCriarVenda.tbTelefone.text
    Email = frmCriarVenda.tbEmail.text
    Documento = frmCriarVenda.tbDocumento.text
    Valor_Total = CDbl(frmCriarVenda.lbTotalValor.Caption)
    Status = 0
    Categoria = Val(Left(frmCriarVenda.coCategoria.text, InStr(frmCriarVenda.coCategoria.text, "-") - 1))
        
    Call Log("Exportando valores para o Banco de Dados...")
    
    'Tratamento de erro ao tentar abrir o Banco
    Call AbrirDatabase
        
    CSQL = "INSERT INTO ordens (cliente, marca, modelo, placa, telefone, email, categoria, documento, valor_total, hora, usuario, status)"
    CSQL = CSQL & "VALUES('"
    CSQL = CSQL & Cliente & "', '" & Marca & "', '" & Modelo & "', '" & Placa & "', '" & Telefone & "', '" & Email & "', '" & Categoria & "', '" & Documento & "', '" & Valor_Total & "', '" & Hora & "', '" & Usuario
    CSQL = CSQL & "', '1')"
    
    resposta = ExecutarRetornoID(CSQL)
    Call FecharDatabase
    ExportaBancoOrdem = resposta
    
End Function
Public Function ExportaBancoOrdemServicos(id_ordem As Integer, id_linha As Integer, id_servico As Integer, valor As Double, desconto As Double, adicional As Double)
    On Error GoTo trata_erro
    
    Call Log("Exportando valores para o Banco de Dados...")
    
    'Tratamento de erro ao tentar abrir o Banco
    Call AbrirDatabase
        
    CSQL = "INSERT INTO ordem_servicos (id_ordem, id_linha, id_servico, valor, desconto, adicional)"
    CSQL = CSQL & "VALUES('"
    CSQL = CSQL & id_ordem & "', '" & id_linha & "', '" & id_servico & "', '" & valor & "', '" & desconto & "', '" & adicional
    CSQL = CSQL & "')"
    
    Call ExecutarQuery(CSQL)
    Call FecharDatabase
    
    Exit Function
    
trata_erro:
    Call Log("Err.Source & " - " & Err.Description")
    MsgBox "Erro ao Exportar dados para o Banco, verifique as informações do Formulário", vbCritical

End Function

