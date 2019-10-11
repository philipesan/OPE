Attribute VB_Name = "mdCriarTabela"
Public Sub CriaTabelasPontos()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela pontos...")
dbtemp.Execute ("CREATE TABLE pontos (id_ponto autoincrement, endereco text(255) NOT NULL, cep text(8) NOT NULL, telefone text(11), gerente int, hr_operacao text(50) NOT NULL, semana text(10) NOT NULL, flag int)")
dbtemp.Execute ("ALTER TABLE pontos ADD CONSTRAINT pontosPk PRIMARY KEY (id_ponto)")    'Adiciona a chave primaria da tabela
dbtemp.Execute ("CREATE INDEX indexpontos ON pontos(id_ponto)")             'Adiciona o indice da tabela
 
Call Log("Tabela pontos criada com sucesso!")

End Sub
Public Sub CriaTabelasStatus()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela status...")
dbtemp.Execute ("Create table status (id_status autoincrement, nome text(45) NOT NULL)")
dbtemp.Execute ("ALTER TABLE status ADD CONSTRAINT statusPk PRIMARY KEY (id_status)")   'Adiciona a chave primaria da tabela
dbtemp.Execute ("CREATE INDEX indexstatus ON status(id_status)")                        'Adiciona o indice da tabela
'Popula a tabela
dbtemp.Execute ("INSERT INTO status(nome) VALUES ('Aguardando')")
dbtemp.Execute ("INSERT INTO status(nome) VALUES ('Em atendimento')")
dbtemp.Execute ("INSERT INTO status(nome) VALUES ('Concluido')")
dbtemp.Execute ("INSERT INTO status(nome) VALUES ('Pago')")
dbtemp.Execute ("INSERT INTO status(nome) VALUES ('Cancelado')")

Call Log("Tabela status criada com sucesso!")

End Sub
Public Sub CriaTabelasCategorias()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela categorias...")
dbtemp.Execute ("Create table categorias (id_categoria autoincrement, nome text(45) NOT NULL, adicional money NOT NULL, flag int)")
dbtemp.Execute ("ALTER TABLE categorias ADD CONSTRAINT categoriasPk PRIMARY KEY (id_categoria)")    'Adiciona a chave primaria da tabela
dbtemp.Execute ("CREATE INDEX indexcategoria ON categorias(id_categoria)")          'Adiciona o indice da tabela

Call Log("Tabela categorias criada com sucesso!")

End Sub

Public Sub CriaTabelasFuncionarios()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela funcionarios...")
dbtemp.Execute ("Create table funcionarios (matricula autoincrement, nome text(50) NOT NULL, cargo int NOT NULL, senha text NOT NULL, flag int)")
dbtemp.Execute ("ALTER TABLE funcionarios ADD CONSTRAINT funcionarioPk PRIMARY KEY (matricula)")   'Adiciona a chave primaria da tabela
dbtemp.Execute ("ALTER TABLE funcionarios ADD CONSTRAINT cargoFk FOREIGN KEY(cargo) REFERENCES cargos(id_cargo)")     'Adiciona a chave primaria da tabela
dbtemp.Execute ("CREATE INDEX indexfuncionarios ON funcionarios(matricula)")           'Adiciona o indice da tabela
senha = UCase(MD5.DigestStrToHexStr("admin"))
dbtemp.Execute ("INSERT INTO funcionarios (nome, cargo, senha, flag) VALUES ('Admin', 1,'" & senha & "', 0)")

Call Log("Tabela funcionarios criada com sucesso!")

End Sub
Public Sub CriaTabelasCargos()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela cargos...")
dbtemp.Execute ("Create table cargos (id_cargo autoincrement, nome text(50) NOT NULL, salario money NOT NULL, acesso_admin int NOT NULL, acesso_rh int NOT NULL, flag int)")
dbtemp.Execute ("ALTER TABLE cargos ADD CONSTRAINT cargoPk PRIMARY KEY (id_cargo)")     'Adiciona a chave primaria da tabela
dbtemp.Execute ("CREATE INDEX indexcargo ON cargos(id_cargo)")             'Adiciona o indice da tabela
dbtemp.Execute ("INSERT INTO cargos (nome, salario, acesso_admin, acesso_rh, flag) VALUES ('Mestre', 0.0, 1, 1, 0)")
Call Log("Tabela cargos criada com sucesso!")

End Sub
Public Sub CriaTabelasServicos()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela servicos...")
dbtemp.Execute ("Create table servicos (id_servico autoincrement, nome text(50), preco money, descricao text(255), flag int)")
dbtemp.Execute ("ALTER TABLE servicos ADD CONSTRAINT servicosPk PRIMARY KEY (id_servico)")  'Adiciona a chave primaria da tabela
dbtemp.Execute ("CREATE INDEX indexservico ON servicos(id_servico)")          'Adiciona o indice da tabela

Call Log("Tabela servicos criada com sucesso!")

End Sub

Public Sub CriaTabelasOrdens()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela ordem")
dbtemp.Execute ("Create table ordens (id_ordem AUTOINCREMENT, cliente text(50), marca text, modelo text, placa text, telefone text(11), email text(30), categoria int, documento text(11), valor_total money, hora text, usuario INT, status INT)")
dbtemp.Execute ("ALTER TABLE ordens ADD CONSTRAINT ordensPk PRIMARY KEY (id_ordem)")                                        'Adiciona a chave primaria da tabela
dbtemp.Execute ("ALTER TABLE ordens ADD CONSTRAINT categoriaFk FOREIGN KEY(categoria) REFERENCES categorias(id_categoria)")  'Adiciona a chave estrangeira
dbtemp.Execute ("ALTER TABLE ordens ADD CONSTRAINT statusFk FOREIGN KEY(status) REFERENCES status(id_status)")  'Adiciona a chave estrangeira
dbtemp.Execute ("ALTER TABLE ordens ADD CONSTRAINT funcionarios_vendaFk FOREIGN KEY(usuario) REFERENCES funcionarios(matricula)")  'Adiciona a chave estrangeira
dbtemp.Execute ("CREATE INDEX indexordens ON ordens(id_ordem)")                                                            'Adiciona o indice da tabela

Call Log("Tabela ordens criada com sucesso!")

End Sub
Public Sub CriaTabelasOrdem_Servicos()

'Adiciona a Tabela e as linhas
Call Log("Criando Tabela ordem_servicos...")
dbtemp.Execute ("Create table ordem_servicos (id_linha INT, id_ordem INT, id_servico INT, valor money, desconto money, adicional money)")
dbtemp.Execute ("ALTER TABLE ordem_servicos ADD CONSTRAINT ordemFk FOREIGN KEY(id_ordem) REFERENCES ordens(id_ordem)")  'Adiciona a chave estrangeira
dbtemp.Execute ("ALTER TABLE ordem_servicos ADD CONSTRAINT servicoFk FOREIGN KEY(id_servico) REFERENCES servicos(id_servico)")  'Adiciona a chave estrangeira

Call Log("Tabela ordem_servicos criada com sucesso!")

End Sub

