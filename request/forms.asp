<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Asp - Formulários</title>
</head>
<body>
  <h1>Formulário com POST - Request.Form</h1>
  <form method='post' action='default.asp'>
    <label for='firstName'>Primeiro Nome: </label>
    <input type='text' id='firstName' name='firstName'> 
    <label for='lastName'>Último Nome: </label>
    <input type='text' id='lastName' name='lastName'>
    <input type='submit' value='Submit'/>
  </form>
  <%
   if request.form <> "" Then
    firstName = request.form("firstName")
    lastName = request.form("lastName")
    response.write "Olá " & firstName & " " & lastName
   end if 
  %>

  <h1>Formulário com Get - Request.QueryString</h1>
  <form method='get' action='default.asp'>
    <label for='getFirstName'>Primeiro Nome: </label>
    <input type='text' id='getFirstName' name='getFirstName'> 
    <label for='getLastName'>Último Nome: </label>
    <input type='text' id='getLastName' name='getLastName'>
    <input type='submit' value='Submit'/>
  </form>
  <%
   if request.QueryString <> "" Then
    getFirstName = request.QueryString("getFirstName")
    getLastName = request.QueryString("getLastName")
    response.write "Olá " & getFirstName & " " & getLastName
   end if 
  %>
</body> 
</html>