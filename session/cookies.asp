<% 
  'Criação de cookies vem antes do HTML, para evitar erros
  if request.form <> "" then
    usuario = request.form("user")
    response.cookies("user") = usuario 'Armazena o cookie chamado "user"
    response.cookies("user").Expires=#Sep, 25, 2021# 'Define data de expiração do cookie, padrão é ao fechar navegador
    message = "Usuário " & request.cookies("user") & " conectado!" 'request.cookie - Pega o conteudo do cookies especificado
  else
    message = "Escreva o nome do usuário!"
  end if
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Asp - Cookies</title>
</head>
<body>
  <h1>Cookies</h1>
  <h2>
    <%
      response.write message
    %>
  </h2>
  <!-- O cookie se mantêm até em outras páginas, não precisa ser na mesma -->
  <form method='post' action='cookies.asp'> 
    <label>Usuário: <input type="text" name="user"></label>
  </form>
</body> 
</html>