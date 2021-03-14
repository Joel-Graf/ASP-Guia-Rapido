# Manual básico de ASP Clássico
Este repositório foi criado com intuito de me familiarizar com a linguagem, deixando registrado alguns dos comandos básicos para consulta posterior.
## Fontes:
[Playlist ASP Clássico](https://www.youtube.com/playlist?list=PLTBN7nX-JQgHd9unHu1Xk0LXeGVBdHYJ8)

### Objetos internos do ASP
#### Application:
	Representa uma aplica��o ASP (Conjusto de p�ginas ASP de um diret�rio virtual no servidor)
	* Objetos globais (global.asa).
	* Aplica��o roda enquanto estiver um browser acessando o servidor.
	* Vari�veis de aplica��o podem ser modificadas por qualquer um.
	* Eventos Application_OnStart e Application_OnEnd.

#### Session:
	Representa uma sess�o aberta com um cliente via browser. Quando uma sess�o se fecha, todas as vari�veis pertencentes a sess�o encerrada s�o perdidas.
	* Se diferencia de aplica��o, pois � individual a cada usu�rio.
	* � apagada ao fechar o browser (padr�o).
	* � mantida atrav�s de cookies (dados no lado do cliente).
	* Possui o met�do ABANDON que encerra a sess�o.
	* Eventos Session_OnStart e Session_OnEnd.

#### Request:
	Representa os dados enviados para a p�gina ASP por um formul�rio ou link do browser do cliente
	* Possui apenas a cole��o de vari�veis enviadas pelo usu�rio atrav�s de um request HTTP.
	* QueryString: Valores enviados por GET.
	* Form: Cont�m todos os campos de um formul�rio enviado por POST.
	* ServerVariables: Cont�m todas as vari�veis CGI do servidor.
	* Cookies: Permite o acesso aos cookies do usu�rio Web.
	* ClienteCertificate: Permite verificar se o cliente possui certifica��o digital (SSL), al�m de acessar dados a respeito dessa certifica��o

#### Response:
	Linhas da p�gina de resposta gerada para o browser do cliente
	* Obs: Tudo que n�o for script � enviado imediatamente ao cliente, script por padr�o n�o fica armazenado em buffer, portanto altera��es em script devem ser declaradas antes do html para terem efeito.
	* Buffer: Indica se a resposta gerada dever� ser enviada automaticamente ou se dever� ser armazenada em um buffer (Deve ser o primeiro comando em uma p�gina ASP).
	* ContentType: � o ContentType do formato MIME da p�gina de resposta.
	* Expires: Usado para definir tempo em que a p�gina ficar� no cache do usu�rio. Em p�ginas est�ticas um tempo maior � vantajoso, em p�ginas din�micas a atualiza��o frequente � prefer�vel.
	* ExpiresAbsolute: Cont�m a data e hora de expira��o da p�gina.
	* Status: Altera as linhas de status gerada pelo servidor.

	* Cole��o de dados - Cookies: Permite criar e alterar os cookies do browser dos usu�rios. Deve ser usado antes de qualquer comando que tamb�m modifique o header http, ou usada com a propriedade buffer = true.
	** Expire: Data e hora de validade de um cookie.
	** Path: Caminho do cookie na m�quina do cliente.
	** Domain: Ser� enviado um cookie para todas as p�ginas contidas no dom�nio especificado.

	* Comandos para enviar informa��es ao browser:
	** AddHeader: Adiciona header HTTP � p�gina gerada.
	** AppendToLog: Nos permite marcar informa��es de acesso ou seguran�a, adicionando um texto ao log do servidor.
	** BinaryWrite: Envia um dado bin�rio para o browser (Deve ser enviado o ContentType em conjunto).
	** Clear: Apaga o conte�do do buffer de resposta do servidor ASP.
	** End: Encerra p�gina ASP.
	** Flush: Envia conte�do do buffer de resposta.
	** Redirect: Redireciona o browser do cliente para outro endere�a (Altera header HTTP, deve ser usado antes de outros conte�dos).
	** Write: Permite o envio de string's para o browser do cliente.


#### Server:
	Representa��o do servidor Web onde a p�gina
	* Tem a propriedade ScriptTimeout, que determina tempo de execu��o do script
	* M�todos:
		** CreateObject: Cria inst�ncia de componente ActiveX, para destruir se atribui o valor constante Nothing.
		** HTMLEncode: Codifica o texto de maneira a n�o interpretar HTML.
		** MapPath: Verifica o caminho f�sico de um diret�rio virtual.
		** URLEnconde: Cofica o texto de maneira correta para URL's.

##### ADO - Conex�o ao Banco de Dados
	� o met�do que veio substituir o DAO e RDO em aplica��es Web, � construido como um objeto ActiveX e abstrai a conex�o ao banco de dados.
	* Recordset: Objeto que representa um conjunto de registros resultantes do processamento de uma instru��o SQL. Existe 3 formas de gerar um Recordset:
	** M�todo Execute de uma Connection: Set RS = DB.Execute("SELECT * FROM Cli")
	** M�todo Execute de um Command: Inst.CommandText = "Select * FROM Cli" : Set Rs = Inst.Execute
	** M�todo Open do Recordset: Set RS = Server.CreateObject("ADODB.Recordset") : Rs.OpenInst

	* Field: Objeto associado a cole��o fields, representa campos de um recordset.

	Objeto Connection: Para criar um objeto connection, devemos ter um datasource ODBC criado previamente, ou criar manualmente atrav�s da string de conex�o.
	Propriedades:
	* CommandTimeout: Tempo m�ximo de espera pela execu��o de um comando.
	* ConnectionString: String de conex�o ao banco de dados ODBC. Pode ser nome do datasource ou os par�metros ODBC. Essa propriedade pode ser omitida caso seja usado o m�todo Open, informando diretamente nele o ConnectionString.
	* Connection.Timeout: Tempo m�ximo para tentar abrir uma conex�o.
	M�todos:
	* Open: Abre uma Connection iniciando uma tentativa de conex�o com o Banco de Dados.
	* Close: Fecha uma conex�o ativa, liberando os recurso do servidor de Banco de Dados.
	* Execute Esse m�todo executa um comando atrav�s de um objeto Connection.
	* BeginTrans: Inicia uma transa��o no Banco de Dados, caso seja poss�vel (Permite Rollback).
	* CommitTrans: Encerra transa��o com sucesso.
	* RollBackTrans: Finaliza transa��o voltando ao estado anterior.

	Objeto Recordset
