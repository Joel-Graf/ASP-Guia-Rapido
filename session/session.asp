<%
  if session("entrada") = "" then 'Caso não exista sessão, inicie uma
    session("entrada") = Time()   'Armazena hora atual na sessão 'entrada'
  end if
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
      response.write "Usuário conectado às " & session("entrada") 'Recupera valor da sessão
    %>
  </h1>
</body> 
</html>