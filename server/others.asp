

<%
  'Renderiza o texto HTML da maneira que foi escrito (Sem interpretar)
  texto = "<br><i>Texto <br>de <strong>teste</strong></i>"
  response.write server.HTMLEncode(texto)
%>
<br><br>