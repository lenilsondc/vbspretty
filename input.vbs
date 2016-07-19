'PARA REALIZAR A MIGRAÇÃO, BASTA EXPORTAR AS PLANILHAS PARA:
'Migracao_Congresso2015 e
'Migracao_CongressoFinanceiro2015
'ALTERE O ID DO EVENTO ABAIXO E O ID DA CATEGORIA DE REMIDO

Dim codEvento, codCategoriaInscricaoRemido
codEvento = 743
codCategoriaInscricaoRemido = 1303

Server.ScriptTimeout = 3600

'Recuperar Inscrições realizadas no local do evento
Set getInscricoesLocalEvento = Server.CreateObject("ADODB.Recordset")
	getInscricoesLocalEvento.ActiveConnection = MM_Conn_BD_STRING
	getInscricoesLocalEvento.Source = "SELECT A.idvisitante, A.departamento, A.categoria, A.tipocracha, A.speaker, A.speakerstatus, A.especialidade, A.nome, A.cracha, A.cpf, A.rg, A.datanasc, A.conselho, A.crm, A.crmuf, A.sexo, A.endereco, A.numero, A.complemento, A.bairro, A.cep, A.cidade, A.uf, A.pais, A.telefone, A.recado, A.celular, A.email, A.valorpre, A.patrocinador, A.cupom, A.pendenciadocs, A.cna, A.sms, A.emitido, A.emitidopor, A.material, A.guardavolume, A.certificado, A.categoriasga, A.situacaosga, A.situacao, B.documento, B.referente, B.valor, B.forma, B.autorizacao, B.vezes, B.datainsc FROM Migracao_Congresso2016 AS A INNER JOIN Migracao_CongressoFinanceiro2016 AS B ON A.idrecibo = B.idrecibo WHERE (B.referente = 'INSCRICAO CONGRESSO') AND a.categoria <> '-' AND a.situacao <> 'CANCELADO' AND codUsuario IS NULL AND codPedido IS NULL /*ORDER BY A.idvisitante*/ "&_
		" UNION ALL "&_
		"SELECT A.idvisitante, A.departamento, A.categoria, A.tipocracha, A.speaker, A.speakerstatus, A.especialidade, A.nome, A.cracha, A.cpf, A.rg, A.datanasc, A.conselho, A.crm, A.crmuf, A.sexo, A.endereco, A.numero, A.complemento, A.bairro, A.cep, A.cidade, A.uf, A.pais, A.telefone, A.recado, A.celular, A.email, A.valorpre, A.patrocinador, A.cupom, A.pendenciadocs, A.cna, A.sms, A.emitido, A.emitidopor, A.material, A.guardavolume, A.certificado, A.categoriasga, A.situacaosga, A.situacao, '' AS documento/*B.documento*/, 'INSCRICAO CONGRESSO' AS referente, A.valor, CASE WHEN A.forma = 'CUPOM' THEN A.forma ELSE 'BOLETO' END AS forma, '' AS autorizacao/*B.autorizacao*/, 1 AS vezes, A.datainsc FROM Migracao_Congresso2016 AS A LEFT OUTER JOIN Migracao_CongressoFinanceiro2016 AS B ON A.idrecibo = B.idrecibo WHERE /*B.referente = 'INSCRICAO CONGRESSO' AND */ a.Forma IN('CORTESIA', 'CUPOM') AND a.categoria <> '-' AND a.situacao <> 'CANCELADO' AND codUsuario IS NULL AND codPedido IS NULL"
	getInscricoesLocalEvento.CursorType = 0
	getInscricoesLocalEvento.CursorLocation = 3
	getInscricoesLocalEvento.LockType = 1
	getInscricoesLocalEvento.Open()


nrDeImportados = 0

ContUsuariosInseridos = 0
LogUsuarioInseridos = "Código dos Usuário Inseridos que fizeram inscrição no local do evento: "

ContPedidosCriados = 0
LogPedidosCriados = "Código dos Pedidos Gerados: "

ContPedidosBaixados = 0
LogPedidosBaixados = "Código dos Pedidos Baixados: "


While (not getInscricoesLocalEvento.EOF)

	nrDeImportados = nrDeImportados + 1

	'Recuperar dados chave para verificar se usuário já está cadastrado na SOCESP
	cpf = "'" & Replace(Trim(AdequaCPF(getInscricoesLocalEvento.Fields.Item("cpf").value)), "'", "''") & "'"
	email = "'" & Replace(Trim(LCase(getInscricoesLocalEvento.Fields.Item("email").value)), "'", "''") & "'"

	'Caso o CPF seja inválido, altera o seu valor de forma que a busca não traga resultados
	if (cpf = "'000.000.000-00'") Then
		cpfQuery = "CPF inválido"
	else
		cpfQuery = Replace(cpf, "'", "")
	end if

	Set ValidaDados = Server.CreateObject("ADODB.Command")
	ValidaDados.ActiveConnection = MM_Conn_BD_STRING
	ValidaDados.CommandText = "dbo.SGA_Checa_Associado"
	ValidaDados.Parameters.Append ValidaDados.CreateParameter("@RETURN_VALUE", 3, 4)
	ValidaDados.Parameters.Append ValidaDados.CreateParameter("@CPF", 200, 1, 20, cpfQuery )
	ValidaDados.Parameters.Append ValidaDados.CreateParameter("@Passaporte", 200, 1, 30, "")
	ValidaDados.Parameters.Append ValidaDados.CreateParameter("@Email", 200, 1, 200, Replace(email, "'", "") )
	ValidaDados.Parameters.Append ValidaDados.CreateParameter("@cod_usuario", 3, 1, 20, 0)
	ValidaDados.CommandType = 4
	ValidaDados.CommandTimeout = 0
	ValidaDados.Prepared = true
	ValidaDados.Execute()

	Resposta = ValidaDados.Parameters.Item("@RETURN_VALUE").Value

	set ValidaDados = Nothing

	'response.Write(Resposta = 1 AND Not isNull(email) AND email <> "")
	'RESPONSE.Write("<br/> " & cpf & " | " & email & " | ")
	'RESPONSE.End()

	'Caso não esteja cadastrado e possua e-mail pois o login é e-mail
	if ( Resposta = 1 AND Not isNull(email) AND email <> "" ) then

		'------------------------
		' Usuário
		'------------------------

		nome = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("nome").value), "'", "''") & "'"

		telefone = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("telefone").value) AND instr(getInscricoesLocalEvento.Fields.Item("telefone").value, "0000-0000") = 0 AND instr(getInscricoesLocalEvento.Fields.Item("telefone").value, "1111-1111") = 0 then
			telefone = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("telefone").value), "'", "''") & "'"
		end if

		celular = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("celular").value) AND instr(getInscricoesLocalEvento.Fields.Item("celular").value, "0000-0000") = 0 AND instr(getInscricoesLocalEvento.Fields.Item("celular").value, "1111-1111") = 0 then
			celular = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("celular").value), "'", "''") & "'"
		end if

		rg = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("rg").value) then
			rg = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("rg").value), "'", "''") & "'"
	 	end if

		pais = "NULL"

		estrangeiro = 0
		if getInscricoesLocalEvento.Fields.Item("pais").value <> "BRASIL" AND getInscricoesLocalEvento.Fields.Item("pais").value <> "BRAISL" AND getInscricoesLocalEvento.Fields.Item("pais").value <> "BRA" AND getInscricoesLocalEvento.Fields.Item("pais").value <> "BRASOL" AND Not isNull(getInscricoesLocalEvento.Fields.Item("pais").value) then
			estrangeiro = 1
			pais = "'BRASIL'"
		else
			if not isNull(getInscricoesLocalEvento.Fields.Item("pais").value) then
				pais = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("pais").value), "'", "''") & "'"
			end if
		end if

		recebeSMS = 0
		if getInscricoesLocalEvento.Fields.Item("sms").value = "SIM" then
			recebeSMS = 1
		end if

		sexo = "NULL"
		if getInscricoesLocalEvento.Fields.Item("sexo").value = "MASCULINO" then
			sexo = "'M'"
		else
			sexo = "'F'"
		end if

		if (instr(cpf, "000.000.000-00") > 0) THen
			cpf = "NULL"
		end if

		if (email = "'00000000@HOTMAIL.COM'") OR (email = "'naopossui@hotmail.com'") Or (email = "XXX@XXXX.COM") Then
			email = "NULL"
		end if

		Set InsertUsuario = Server.CreateObject ("ADODB.Command")
			InsertUsuario.ActiveConnection = MM_Conn_BD_STRING
			InsertUsuario.CommandText = "INSERT INTO User_Usuarios ( NomeCompleto, email, Telefone, Celular, CPF, RG, Estrangeiro, recebeSMS, Sexo, cod_TipoUsuario ) VALUES (" & nome & ", " & email & ", " & telefone & ", " & celular & ", " & cpf & ", " & rg & ", " & estrangeiro & ", " & recebeSMS & ", " & sexo & ", 8)"
			InsertUsuario.CommandType = 1
			InsertUsuario.CommandTimeout = 0
			InsertUsuario.Prepared = true
			response.Write(InsertUsuario.CommandText & "<br />")
			InsertUsuario.Execute()
		Set InsertUsuario = Nothing

		'------------------------
		' Associado
		'------------------------

		'Recuperar código do usuário inserido
		codUsuario = exQu("SELECT TOP(1) cod_Usuario FROM User_Usuarios ORDER BY cod_Usuario DESC", "cod_Usuario")
		codAssociado = codUsuario


		codTipoConselho = 15 'Sem Conselho
		codPendencias = 9 'Sem pedência
		DataTituloEspe = "NULL"
		DataFiliacao = "NULL"

		DataNascimento = "NULL"
		if( getInscricoesLocalEvento.Fields.Item("dataNasc").value <> "" AND Not isNull(getInscricoesLocalEvento.Fields.Item("dataNasc").value) )then
			if( isDate(getInscricoesLocalEvento.Fields.Item("dataNasc").value) )then
				DataNascimento = "CONVERT(datetime, '" & getInscricoesLocalEvento.Fields.Item("dataNasc").value & "', 103)"
			end if
		end if

		UFConselho = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("crmuf").value) then
			UFConselho = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("crmuf").value), "'", "''") & "'"
		end if

		NumeroConselho = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("crm").value) then
			NumeroConselho = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("crm").value), "'", "''") & "'"
		end if

		codDepartamento = 32
		if not isNull(getInscricoesLocalEvento.Fields.Item("departamento").value) then
			codDepartamento = exQu("SELECT cod_Departamento FROM SGA_AssociadosDepartamento WHERE departamento = '" & getInscricoesLocalEvento.Fields.Item("departamento").value & "'", "cod_Departamento")
		end if

		codEspecialidade = 1
		if not isNull(getInscricoesLocalEvento.Fields.Item("especialidade").value) AND getInscricoesLocalEvento.Fields.Item("especialidade").value <> "NAO INFORMADA" then
			nomeEspecialidade = getInscricoesLocalEvento.Fields.Item("especialidade").value
			if (nomeEspecialidade = "NEFROLOGISTA") Then
				nomeEspecialidade = "Nefrologia"
			end if

			codEspecialidade = exQu("SELECT cod_Especialidade FROM SGA_AssociadosEspecialidade WHERE especialidade = '" & nomeEspecialidade & "'", "cod_Especialidade")
		end if

		codResidencia = 3

		Set InsereAssociado = Server.CreateObject ("ADODB.Command")
			InsereAssociado.ActiveConnection = MM_Conn_BD_STRING
			InsereAssociado.CommandText = "INSERT INTO SGA_Associados ( cod_associado, cod_Usuario, cod_TipoConselho, cod_Pendencias, DataTituloEspe, DataFilicao, DataNascimento, UF_Conselho, Numero_Conselho, cod_Departamento, cod_Especialidade, cod_Residencia ) VALUES (" & codAssociado & ", " & codUsuario & ", " & codTipoConselho & ", " & codPendencias & ", " & DataTituloEspe & ", " & DataFiliacao & ", " & DataNascimento & ", " & UFConselho & ", " & NumeroConselho & ", " & codDepartamento & ", " & codEspecialidade & ", " & codResidencia & ")"
			InsereAssociado.CommandType = 1
			InsereAssociado.CommandTimeout = 0
			InsereAssociado.Prepared = true
			response.Write(InsereAssociado.CommandText & "<br />")
			InsereAssociado.Execute()
		Set InsereAssociado = Nothing

		'------------------------
		' Categoria associado
		'------------------------

		'Definir Categoria do associado
		if getInscricoesLocalEvento.Fields.Item("departamento").value = "MÉDICO" OR getInscricoesLocalEvento.Fields.Item("departamento").value = "MEDICINA ACADÊMICO" OR getInscricoesLocalEvento.Fields.Item("departamento").value = "RESIDENTES" then
			codCategoriaAssociado = 7 'Cadastrado Área Médica
		else
			codCategoriaAssociado = 14 'Cadastro Departamentos
		end if

		ExUp("UPDATE SGA_Associados SET cod_Categorias = " & codCategoriaAssociado & " WHERE cod_associado = " & codAssociado)

		'------------------------
		' Endereço
		'------------------------

		endereco = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("endereco").value) then
			endereco = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("endereco").value), "'", "''") & "'"
		end if

		numeroEnd = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("numero").value) then
			numeroEnd = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("numero").value), "'", "''") & "'"
		end if

		complemento = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("complemento").value) then
			complemento = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("complemento").value), "'", "''") & "'"
		end if

		bairro = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("bairro").value) then
			bairro = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("bairro").value), "'", "''") & "'"
		end if

		cidade = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("cidade").value) then
			cidade = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("cidade").value), "'", "''") & "'"
		end if

		uf = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("uf").value) then
			uf = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("uf").value), "'", "''") & "'"
		end if

		cep = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("cep").value) then
			cep = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("cep").value), "'", "''") & "'"
		end if

		if (not isNull(getInscricoesLocalEvento.Fields.Item("cep").value)) AND getInscricoesLocalEvento.Fields.Item("cep").value <> "00000-000" AND getInscricoesLocalEvento.Fields.Item("cep").value <> "11111-111" Then

			Set InsereUsuarioEndereco = Server.CreateObject ("ADODB.Command")
				InsereUsuarioEndereco.ActiveConnection = MM_Conn_BD_STRING
				InsereUsuarioEndereco.CommandText = "INSERT INTO User_Enderecos ( cod_Usuario, Enderecos, NumeroEnd, Complemento, Bairro, Cidade, UF, CEP, Pais, Entrega, Divulgacao ) VALUES ("& codUsuario & ", " & endereco & ", " & numeroEnd & ", " & complemento & ", " & bairro & ", " & cidade & ", " & uf & ", " & cep & ", " & pais & ", 1, 0)"
				InsereUsuarioEndereco.CommandType = 1
				InsereUsuarioEndereco.CommandTimeout = 0
				InsereUsuarioEndereco.Prepared = true
				response.Write(InsereUsuarioEndereco.CommandText & "<br />")
				InsereUsuarioEndereco.Execute()
			Set InsereUsuarioEndereco = Nothing

		end if

		'------------------------
		' Permissões
		'------------------------

		Set InserePermissao1 = Server.CreateObject("ADODB.Command")
			InserePermissao1.ActiveConnection = MM_Conn_BD_STRING
			InserePermissao1.CommandText = "INSERT INTO dbo.User_Permissoes ( cod_usuario, AcessoNivel ) VALUES (" & codUsuario & ", 1)"
			InserePermissao1.CommandType = 1
			InserePermissao1.CommandTimeout = 0
			InserePermissao1.Prepared = true
			response.Write(InserePermissao1.CommandText & "<br />")
			InserePermissao1.Execute()
		Set InserePermissao1 = Nothing

		Set InserePermissao2 = Server.CreateObject("ADODB.Command")
			InserePermissao2.ActiveConnection = MM_Conn_BD_STRING
			InserePermissao2.CommandText = "INSERT INTO dbo.User_Permissoes ( cod_usuario, AcessoNivel ) VALUES (" & codUsuario & ", 2)"
			InserePermissao2.CommandType = 1
			InserePermissao2.CommandTimeout = 0
			InserePermissao2.Prepared = true
			response.Write(InserePermissao2.CommandText & "<br />")
			InserePermissao2.Execute()
		Set InserePermissao2 = Nothing


		'------------------------
		' Informações do médico
		'------------------------

		nomeCracha = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("cracha").Value) then
			nomeCracha =   "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("cracha").Value),"'","''") & "'"
		end if

		Set InsereInfoMedico = Server.CreateObject("ADODB.Command")
			InsereInfoMedico.ActiveConnection = MM_Conn_BD_STRING
			InsereInfoMedico.CommandText = "INSERT INTO User_infoMedico( NomeCracha, cod_usuario ) VALUES ( " & nomeCracha & ", " & codUsuario & ")"
			InsereInfoMedico.CommandType = 1
			InsereInfoMedico.CommandTimeout = 0
			InsereInfoMedico.Prepared = true
			response.Write(InsereInfoMedico.CommandText & "<br />")
			InsereInfoMedico.Execute()
		Set InsereInfoMedico = Nothing


		'------------------------
		' Login
		'------------------------

		login = email
		senha = CriaSenha(3)

		Set InsereLogin = Server.CreateObject("ADODB.Command")
			InsereLogin.ActiveConnection = MM_Conn_BD_STRING
			InsereLogin.CommandText = "INSERT INTO dbo.User_Login ( cod_usuario, Login, Senha ) VALUES ( " & codUsuario & ", " & login & ", '" & senha & "')"
			InsereLogin.CommandType = 1
			InsereLogin.CommandTimeout = 0
			InsereLogin.Prepared = true
			response.Write(InsereLogin.CommandText & "<br />")
			InsereLogin.Execute()
		Set InsereLogin = Nothing

		'Insere o código do usuário gerado
		ExUp("UPDATE Migracao_Congresso2015 SET codUsuario = " & codUsuario & " WHERE idvisitante = " & getInscricoesLocalEvento.Fields.Item("idvisitante").Value)

		LogUsuarioInseridos = LogUsuarioInseridos & codUsuario & ", "
		ContUsuariosInseridos = ContUsuariosInseridos + 1

	end if 'if Resposta = 1 then

	if not isNull(cpf) AND cpf <> "" and cpf <> "'000.000.000-00'" then
		query = "SELECT cod_Usuario, NomeCompleto, email, CPF, Estrangeiro FROM User_Usuarios WHERE CPF = " & cpf

		Set getByCpf = Server.CreateObject("ADODB.Recordset")
		getByCpf.ActiveConnection = MM_Conn_BD_STRING
		getByCpf.Source = "SELECT COUNT(cod_Usuario) AS Quantidade FROM User_Usuarios WHERE CPF = " & cpf
		getByCpf.CursorType = 0
		getByCpf.CursorLocation = 3
		getByCpf.LockType = 1
		getByCpf.Open()

		if (getByCpf.Fields.Item("Quantidade").value = 0) Then
			Response.Write("Usuário não encontrado pelo CPF<br/>")
			query = "SELECT cod_Usuario, NomeCompleto, email, CPF, Estrangeiro FROM User_Usuarios WHERE email = " & email
		end if

	else
		query = "SELECT cod_Usuario, NomeCompleto, email, CPF, Estrangeiro FROM User_Usuarios WHERE email = " & email
	end if

	'recuperar dados do usuário para realizar inscrição
	Set getDadosUsuario = Server.CreateObject("ADODB.Recordset")
		getDadosUsuario.ActiveConnection = MM_Conn_BD_STRING
		getDadosUsuario.Source = query
		getDadosUsuario.CursorType = 0
		getDadosUsuario.CursorLocation = 3
		getDadosUsuario.LockType = 1
		Response.Write(query & "<br/>")
		getDadosUsuario.Open()

		codUsuario = getDadosUsuario.Fields.Item("cod_Usuario").value
		NomeCliente = "'" & Replace(Trim(getDadosUsuario.Fields.Item("NomeCompleto").value), "'", "''") & "'"

		getDadosUsuario.Close()
	Set getDadosUsuario = Nothing


	'Verificar se não está inscrito
	estaInscrito = exQu("SELECT COUNT(cod_Inscricao) AS Inscricoes FROM Cursos_Inscricao WHERE cod_Eventos = " & codEvento& " AND cod_usuario = " & codUsuario, "Inscricoes")

	estaInscritoEmAberto = exQu("SELECT COUNT(cod_Inscricao) AS Inscricoes FROM Cursos_Inscricao WHERE cod_StatusInscrito NOT IN(4, 5) AND cod_Eventos = " & codEvento & " AND cod_usuario = " & codUsuario, "Inscricoes")

	'Fazer inscrição caso não esteja inscrito
	if estaInscrito = 0 OR estaInscritoEmAberto > 0 then

		'------------------------
		' Pedido
		'------------------------

		DataInicioVenda = "NULL"
		if( getInscricoesLocalEvento.Fields.Item("datainsc").value <> "" AND Not isnull(getInscricoesLocalEvento.Fields.Item("datainsc").value) )then
			if( isDate(getInscricoesLocalEvento.Fields.Item("datainsc").value) )then
				DataInicioVenda = "CONVERT(datetime, '" & getInscricoesLocalEvento.Fields.Item("datainsc").value & "', 102)"
			end if
		end if
		'Demais datas tem o mesmo valor da data de início da venda
		DataFinalVenda  = DataInicioVenda
		dataVencimento  = DataInicioVenda
		dataPagamento   = DataInicioVenda
		DataFinalizacao = DataInicioVenda

		codStatusVenda = 8 'Pedido Pago

		'Recuperar Informações para realizar inscrição
		valorTotal = Trim(getInscricoesLocalEvento.Fields.Item("valor").value)

		if (valorTotal = "-") Then
			valorTotal = "0"
		end if

		valorTotal = Replace(valorTotal, ".", "")
		valorTotal = Replace(valorTotal, ",", ".")
		valorTotal = "'" & Replace(valorTotal, "'", "''") & "'"

		codTipoPedido = 2 'Inscrição

		formaPagamento = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("forma").value), "'", "''") & "'"
		codFormaPagamento = "NULL"
		if not isNull(formaPagamento) then
			codFormaPagamento = exQu("SELECT cod_FormaPagamento FROM LJ_FormaPagamento WHERE LJ_FormaPagamento = " & formaPagamento, "cod_FormaPagamento")
		else
			codFormaPagamento = 6 'Dinheiro
		end if

		if (isNull(getInscricoesLocalEvento.Fields.Item("vezes").value)) Then
			Parcelas = "1"
		else
			Parcelas = Replace(Trim(getInscricoesLocalEvento.Fields.Item("vezes").value), "'", "''")
		end if
		NumeroRecibo =  "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("documento").value), "'", "''") & "'"
		NumeroAutorizacao = "NULL"
		if not isNull(getInscricoesLocalEvento.Fields.Item("autorizacao").value) then
			NumeroAutorizacao = "'" & Replace(Trim(getInscricoesLocalEvento.Fields.Item("autorizacao").value), "'", "''") & "'"
		end if

		'Se possuir uma inscrição em aberto faz o update
		codPedido = exQu("SELECT cod_Pedido FROM Cursos_Inscricao WHERE cod_StatusInscrito NOT IN(4, 5) AND cod_Eventos = " & codEvento & " AND cod_usuario = " & codUsuario, "cod_Pedido")
		if estaInscritoEmAberto > 0 And codPedido <> "" then

			'Realiza a atualização do Pedido
			ExUp("UPDATE LJ_Pedido SET NomeCliente =" & NomeCliente & ", cod_usuario = " & codUsuario & ", DataInicioVenda = " & DataInicioVenda & ", DataFinalVenda = " & DataFinalVenda & ", dataVencimento = " & dataVencimento & ", dataPagamento = " & dataPagamento & ", DataFinalizacao = " & DataFinalizacao & ", cod_StatusVenda = " & codStatusVenda & ", ValorTotal = " & valorTotal & ", codTipoPedido = " & codTipoPedido & ", cod_FormaPagamento = " & codFormaPagamento & ", Parcelas = " & Parcelas & ", NumeroRecibo = " & NumeroRecibo & ", NumeroAutorizacao = " & NumeroAutorizacao & " WHERE cod_Vendas = " & codPedido)

			'Realiza a atualização do Registro de Migração para constar o Id do Pedido
			ExUp("UPDATE Migracao_Congresso2015 SET codPedido = " & codPedido & " WHERE idvisitante = " & getInscricoesLocalEvento.Fields.Item("idvisitante").Value)
		else
			Set gerarPedido = Server.CreateObject("ADODB.Command")
			gerarPedido.ActiveConnection = MM_Conn_BD_STRING
			gerarPedido.CommandText = "INSERT INTO LJ_Pedido( NomeCliente, cod_usuario, DataInicioVenda, DataFinalVenda, dataVencimento, dataPagamento, DataFinalizacao, cod_StatusVenda, ValorTotal,  codTipoPedido, cod_FormaPagamento, Parcelas, NumeroRecibo, NumeroAutorizacao ) VALUES (" & NomeCliente & ", " & codUsuario & ", " & DataInicioVenda & ", " & DataFinalVenda & ", " & dataVencimento & ", " & dataPagamento & ", " & DataFinalizacao & ", " & codStatusVenda & ", " & valorTotal & ", " & codTipoPedido & ", " & codFormaPagamento & ", " & Parcelas & ", " & NumeroRecibo & ", " & NumeroAutorizacao & " )"
			gerarPedido.CommandType = 1
			gerarPedido.CommandTimeout = 0
			gerarPedido.Prepared = true
			response.Write(gerarPedido.CommandText & "<br />")
			gerarPedido.Execute()
			Set gerarPedido = Nothing

			'Recupera o código do pedido
			codPedido = ExQu("SELECT MAX(cod_Vendas) AS ultimoID FROM LJ_Pedido", "ultimoID")

			'Realiza a atualização do Registro de Migração para constar o Id do Pedido
			ExUp("UPDATE Migracao_Congresso2015 SET codPedido = " & codPedido & " WHERE idvisitante = " & getInscricoesLocalEvento.Fields.Item("idvisitante").Value)

			LogPedidosCriados = LogPedidosCriados & codPedido & ", "
			ContPedidosCriados = ContPedidosCriados + 1
		end if

		'------------------------
		' Inscrição
		'------------------------

		categoriaInscricao = Replace(Trim(getInscricoesLocalEvento.Fields.Item("categoria").value), "'", "''")
		CodCategoriaInscricao = "NULL"
		if not isNull(categoriaInscricao) AND categoriaInscricao <> "" AND categoriaInscricao <> "'REMIDO SBC'" then

				If categoriaInscricao = "ACADÊMICO - ÁREA DEPARTAMENTO" Then
					CodCategoriaInscricao = 1295
				ElseIf categoriaInscricao = "BIOLOGIA, BIOMEDICINA, BIOMOLECULAR, QUÍMICO E TECNÓLOGO NA ÁREA DE SAÚDE" Then
					CodCategoriaInscricao = 1302
				ElseIf categoriaInscricao = "ESTAGIÁRIO / PÓS-GRADUAÇÃO - ÁREA DEPARTAMENTO" Then
					CodCategoriaInscricao = 1294
				ElseIf categoriaInscricao = "NÃO SÓCIO - OUTROS NÃO CARDIOLOGISTA 26/05" Then
					CodCategoriaInscricao = 1307
				ElseIf categoriaInscricao = "NÃO SÓCIO - OUTROS NÃO CARDIOLOGISTA 27/05" Then
					CodCategoriaInscricao = 1308
				ElseIf categoriaInscricao = "NÃO SÓCIO - OUTROS NÃO CARDIOLOGISTA 28/05" Then
					CodCategoriaInscricao = 1309
				ElseIf categoriaInscricao = "NÃO SÓCIO / NÃO QUITE - ÁREA DEPARTAMENTO" Then
					CodCategoriaInscricao = 1293
				ElseIf categoriaInscricao = "NÃO SÓCIO / NÃO QUITE - ÁREA MÉDICA" Then
					CodCategoriaInscricao = 1287
				ElseIf categoriaInscricao = "NÃO SÓCIO / NÃO QUITE - ODONTOLOGIA, PSICOLOGIA E SERVIÇO SOCIAL (01 DIA)" Then
					CodCategoriaInscricao = 1299
				ElseIf categoriaInscricao = "RESIDENTE / ESTAGIÁRIO / PÓS-GRADUAÇÃO - ÁREA MÉDICA" Then
					CodCategoriaInscricao = 1288
				ElseIf categoriaInscricao = "SÓCIO QUITE - OUTROS NÃO CARDIOLOGISTA 26/05" Then
					CodCategoriaInscricao = 1304
				ElseIf categoriaInscricao = "SÓCIO QUITE - OUTROS NÃO CARDIOLOGISTA 27/05" Then
					CodCategoriaInscricao = 1305
				ElseIf categoriaInscricao = "SÓCIO QUITE - OUTROS NÃO CARDIOLOGISTA 28/05" Then
					CodCategoriaInscricao = 1306
				ElseIf categoriaInscricao = "SÓCIO QUITE 2016 - ÁREA DEPARTAMENTO" Then
					CodCategoriaInscricao = 1296
				ElseIf categoriaInscricao = "SÓCIO QUITE 2016 - ÁREA MÉDICA" Then
					CodCategoriaInscricao = 1290
				ElseIf categoriaInscricao = "SÓCIO QUITE 2016 - ODONTOLOGIA, PSICOLOGIA E SERVIÇO SOCIAL (01 DIA)" Then
					CodCategoriaInscricao = 1300
				End If
		else
			CodCategoriaInscricao = codCategoriaInscricaoRemido 'Remido
		end if

		codTipoInscricao = 4 'Transferência de inscrição.

		autorizaFornecimentoDados = 0

		codStatusInscricao = 4 'Inscrição Paga

		ValorCurso = valorTotal

		'Caso a inscrição esteja em aberto faz a atualização
		if estaInscritoEmAberto > 0 then
			codInscricao = exQu("SELECT TOP(1) cod_Inscricao AS Inscricoes FROM Cursos_Inscricao WHERE cod_StatusInscrito NOT IN(4, 5) AND cod_Eventos = " & codEvento & " AND cod_usuario = " & codUsuario, "Inscricoes")
			'Realiza a atualização da Inscrição
			ExUp("UPDATE Cursos_Inscricao SET cod_usuario = " & codUsuario & ", cod_eventos = " & codEvento & ", Cod_CategoriaInscricao = " & CodCategoriaInscricao & ", cod_TipoInscricao = " & codTipoInscricao & ", autorizaFornecimentoDados = " & autorizaFornecimentoDados & ", cod_Pedido = " & codPedido & ", cod_StatusInscrito = " & codStatusInscricao & ", ValorCurso = " & ValorCurso & " WHERE cod_Inscricao = " & codInscricao)
		else
			Set gerarInscricao = Server.CreateObject("ADODB.Command")
			gerarInscricao.ActiveConnection = MM_Conn_BD_STRING
			gerarInscricao.CommandText = "INSERT INTO Cursos_Inscricao( cod_usuario, cod_eventos, Cod_CategoriaInscricao, cod_TipoInscricao, autorizaFornecimentoDados, cod_Pedido, cod_StatusInscrito, ValorCurso ) VALUES ( "& codUsuario & ", " & codEvento & ", " & CodCategoriaInscricao & ", " & codTipoInscricao & ", " & autorizaFornecimentoDados & ", " & codPedido & ", " & codStatusInscricao & ", " & ValorCurso & " )"
			gerarInscricao.CommandType = 1
			gerarInscricao.CommandTimeout = 0
			gerarInscricao.Prepared = true
			response.Write(gerarInscricao.CommandText & "<br />")
		    gerarInscricao.Execute()
			Set gerarInscricao = Nothing
		end if

		'------------------------
		' Resposta do CNA
		'------------------------
		possuiRespostaCna = exQu("SELECT COUNT(cod_resposta) AS Respostas FROM Cursos_RespostasCNA WHERE cod_usuario = " & codUsuario & " AND cod_Evento = " & codEvento, "Respostas")

		'Caso ainda não possua a resposta do CNA
		if (possuiRespostaCna = 0) Then
			respostaCna = 0
			if getInscricoesLocalEvento.Fields.Item("cna").value = "SIM" then
				respostaCna = 1
			end if

			dataResposta = DataInicioVenda

			Set gerarRespostaCNA = Server.CreateObject("ADODB.Command")
				gerarRespostaCNA.ActiveConnection = MM_Conn_BD_STRING
				gerarRespostaCNA.CommandText = "INSERT INTO Cursos_RespostasCNA( cod_usuario, cod_evento, resposta, data_resposta ) VALUES ( " & codUsuario & ", " & codEvento & ", " & respostaCna & ", " & dataResposta & " )"
				gerarRespostaCNA.CommandType = 1
				gerarRespostaCNA.CommandTimeout = 0
				gerarRespostaCNA.Prepared = true
				response.Write(gerarRespostaCNA.CommandText & "<br />")
				gerarRespostaCNA.Execute()
			Set gerarRespostaCNA = Nothing
		end if
	end if 'if estaInscrito = 0 then

  'Response.End()
  Response.Flush()
  getInscricoesLocalEvento.MoveNext()
Wend

  Acao 21490, "Cong. 2016 - Importação Inscrições no Local - Usuários Cadastrados ", "Importação: " & ContUsuariosInseridos & "Usuários Cadastrados", LogUsuarioInseridos

  Acao 21490, "Cong. 2016 - Importação Inscrições no Local - Pedidos Criados", "Importação: " & ContPedidosCriados & " Pedidos Criados: ", LogPedidosCriados

  Acao 21490, "Cong. 2016 - Importação Inscrições no Local - Pedidos Baixados", "Importação: " & ContPedidosBaixados & " Pedidos Baixados: ", LogPedidosBaixados
