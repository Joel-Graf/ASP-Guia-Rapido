# Manual básico de ASP Clássico
Este repositório foi criado com intuito de me familiarizar com a linguagem, deixando registrado alguns dos comandos básicos para consulta posterior.
## Fontes:
[Playlist ASP Clássico](https://www.youtube.com/playlist?list=PLTBN7nX-JQgHd9unHu1Xk0LXeGVBdHYJ8)

### Objetos internos do ASP
#### Application:
	Representa uma aplicação ASP (Conjusto de páginas ASP de um diretório virtual no servidor)
	* Objetos globais (global.asa).
	* Aplicação roda enquanto estiver um browser acessando o servidor.
	* Variáveis de aplicação podem ser modificadas por qualquer um.
	* Eventos Application_OnStart e Application_OnEnd.

#### Session:
	Representa uma sessão aberta com um cliente via browser. Quando uma sessão se fecha, todas as variáveis pertencentes a sessão encerrada são perdidas.
	* Se diferencia de aplicação, pois é individual a cada usuário.
	* É apagada ao fechar o browser (padrão).
	* É mantida através de cookies (dados no lado do cliente).
	* Possui o metódo ABANDON que encerra a sessão.
	* Eventos Session_OnStart e Session_OnEnd.

#### Request:
	Representa os dados enviados para a página ASP por um formulário ou link do browser do cliente
	* Possui apenas a coleção de variáveis enviadas pelo usuário através de um request HTTP.
	* QueryString: Valores enviados por GET.
	* Form: Contêm todos os campos de um formulário enviado por POST.
	* ServerVariables: Contém todas as variáveis CGI do servidor.
	* Cookies: Permite o acesso aos cookies do usuário Web.
	* ClienteCertificate: Permite verificar se o cliente possui certificação digital (SSL), além de acessar dados a respeito dessa certificação

#### Response:
	Linhas da página de resposta gerada para o browser do cliente
	* Obs: Tudo que não for script é enviado imediatamente ao cliente, script por padrão não fica armazenado em buffer, portanto alterações em script devem ser declaradas antes do html para terem efeito.
	* Buffer: Indica se a resposta gerada deverá ser enviada automaticamente ou se deverá ser armazenada em um buffer (Deve ser o primeiro comando em uma página ASP).
	* ContentType: É o ContentType do formato MIME da página de resposta.
	* Expires: Usado para definir tempo em que a página ficará no cache do usuário. Em páginas estáticas um tempo maior é vantajoso, em páginas dinâmicas a atualização frequente é preferível.
	* ExpiresAbsolute: Contém a data e hora de expiração da página.
	* Status: Altera as linhas de status gerada pelo servidor.

	* Coleção de dados - Cookies: Permite criar e alterar os cookies do browser dos usuários. Deve ser usado antes de qualquer comando que também modifique o header http, ou usada com a propriedade buffer = true.
	** Expire: Data e hora de validade de um cookie.
	** Path: Caminho do cookie na máquina do cliente.
	** Domain: Será enviado um cookie para todas as páginas contidas no domínio especificado.

	* Comandos para enviar informações ao browser:
	** AddHeader: Adiciona header HTTP à página gerada.
	** AppendToLog: Nos permite marcar informações de acesso ou segurança, adicionando um texto ao log do servidor.
	** BinaryWrite: Envia um dado binário para o browser (Deve ser enviado o ContentType em conjunto).
	** Clear: Apaga o conteúdo do buffer de resposta do servidor ASP.
	** End: Encerra página ASP.
	** Flush: Envia conteúdo do buffer de resposta.
	** Redirect: Redireciona o browser do cliente para outro endereça (Altera header HTTP, deve ser usado antes de outros conteúdos).
	** Write: Permite o envio de string's para o browser do cliente.


#### Server:
	Representação do servidor Web onde a página
	* Tem a propriedade ScriptTimeout, que determina tempo de execução do script
	* Métodos:
		** CreateObject: Cria instância de componente ActiveX, para destruir se atribui o valor constante Nothing.
		** HTMLEncode: Codifica o texto de maneira a não interpretar HTML.
		** MapPath: Verifica o caminho físico de um diretório virtual.
		** URLEnconde: Cofica o texto de maneira correta para URL's.

##### ADO - Conexão ao Banco de Dados
	É o metódo que veio substituir o DAO e RDO em aplicações Web, é construido como um objeto ActiveX e abstrai a conexão ao banco de dados.
	* Recordset: Objeto que representa um conjunto de registros resultantes do processamento de uma instrução SQL. Existe 3 formas de gerar um Recordset:
	** Método Execute de uma Connection: Set RS = DB.Execute("SELECT * FROM Cli")
	** Método Execute de um Command: Inst.CommandText = "Select * FROM Cli" : Set Rs = Inst.Execute
	** Método Open do Recordset: Set RS = Server.CreateObject("ADODB.Recordset") : Rs.OpenInst

	* Field: Objeto associado a coleção fields, representa campos de um recordset.

	Objeto Connection: Para criar um objeto connection, devemos ter um datasource ODBC criado previamente, ou criar manualmente através da string de conexão.
	Propriedades:
	* CommandTimeout: Tempo máximo de espera pela execução de um comando.
	* ConnectionString: String de conexão ao banco de dados ODBC. Pode ser nome do datasource ou os parâmetros ODBC. Essa propriedade pode ser omitida caso seja usado o método Open, informando diretamente nele o ConnectionString.
	* Connection.Timeout: Tempo máximo para tentar abrir uma conexão.
	Métodos:
	* Open: Abre uma Connection iniciando uma tentativa de conexão com o Banco de Dados.
	* Close: Fecha uma conexão ativa, liberando os recurso do servidor de Banco de Dados.
	* Execute Esse método executa um comando através de um objeto Connection.
	* BeginTrans: Inicia uma transação no Banco de Dados, caso seja possível (Permite Rollback).
	* CommitTrans: Encerra transação com sucesso.
	* RollBackTrans: Finaliza transação voltando ao estado anterior.

	Objeto Recordset
