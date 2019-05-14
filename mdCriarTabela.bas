Attribute VB_Name = "mdCriarTabela"
Public Sub CriaTabelasPontos()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela pontos...")
dbtemp.Execute ("Create table pontos (id_ponto autoincrement, endereco text(255), cep text(8), telefone text(11), gerente int, hr_operacao text(50), semana text(10))")
dbtemp.Execute ("Create index chaveponto on pontos(id_ponto)")           'Adiciona o indice da tabela
Call Log("Tabela pontos criada com sucesso!")

End Sub
Public Sub CriaTabelasStatus()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela status...")
dbtemp.Execute ("Create table status (id_status autoincrement, nome text(45))")
dbtemp.Execute ("Create index chavestatus on status(id_status)")           'Adiciona o indice da tabela
Call Log("Tabela status criada com sucesso!")

End Sub
Public Sub CriaTabelasCategorias()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela categorias...")
dbtemp.Execute ("Create table categorias (id_categoria autoincrement, nome text(45), adicional money)")
dbtemp.Execute ("Create index chavecategoria on categorias(id_categoria)")           'Adiciona o indice da tabela
Call Log("Tabela categorias criada com sucesso!")

End Sub

Public Sub CriaTabelasFuncionarios()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela funcionarios...")
dbtemp.Execute ("Create table funcionarios (matricula autoincrement, nome text(50), cargo int)")
dbtemp.Execute ("Create index chavefuncionario on funcionarios(matricula)")           'Adiciona o indice da tabela
Call Log("Tabela funcionarios criada com sucesso!")

End Sub
Public Sub CriaTabelasCargos()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela cargos...")
dbtemp.Execute ("Create table cargos (id_cargo autoincrement, nome text(50), salario money, acesso_admin int, acesso_rh int)")
dbtemp.Execute ("Create index chavecargos on cargos(id_cargo)")           'Adiciona o indice da tabela
Call Log("Tabela cargos criada com sucesso!")

End Sub
