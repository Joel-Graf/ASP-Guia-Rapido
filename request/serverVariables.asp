<%
  'request.serverVariables - Acessa informações do servidor e usuário
  navegador = request.serverVariables("http_user_agent")
  servidor = request.serverVariables("server_software")
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Asp - Sessão</title>
</head>
<body>
  <h1>
    <%
      response.write navegador
      response.write servidor
    %>
  </h1>
</body> 
</html>