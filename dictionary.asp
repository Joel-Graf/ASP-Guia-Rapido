<%
  'Data Structure - Dictionary
  'Basicamente um array com chave textual
  set d = server.createObject("Scripting.Dictionary")

  d.Add "SP", "São Paulo"
  d.Add "RJ", "Rio de Janeiro"
  d.Add "MG", "Minas Gerais"
  d.Add "ES", "Espírito Santo" 

  response.write("SP é a sigla de " & d.item("SP"))
  response.write("MG é a sigla de " & d.item("MG"))

  'Propriedades do Dicionary
  total = d.count 'Conta os itens
  response.write "Existem " & total & " Itens!"
%>