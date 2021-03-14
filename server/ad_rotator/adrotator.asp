<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ASP - Adrotator</title>
  <style> 
  </style>
</head>
<body>
  <%
    set adrotator = server.createObject("MSWC.AdRotator")
    response.write(adrotator.getadvertisement("banners.txt"))
  %>
</body> 
</html>