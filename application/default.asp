<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Asp - Global.asa</title>
</head>
<body>
  <h1>Global.asa</h1>
  <p>É um arquivo para se armazenar variáveis e métodos globais</p>
  <%
    'Variáveis globais são usadas da mesma forma que locais
    response.write "Horário de Conexão: " & Session("iniciou")
  %>
</body> 
</html>