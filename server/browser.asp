<%
  set Navegador = Server.createObject("MSWC.BrowserType")
  response.write "Navegador: " & navegador.browser & "<br>"
  response.write "Versão Navegador: " & navegador.version & "<br>"
  response.write "Sistema Operacial: " & navegador.plataform & "<br>"
  response.write "Suporta Frames?: " & navegador.frames & "<br>"
  response.write "Suporta Tabelas?: " & navegador.tables & "<br>"
  response.write "Suporta Sons?: " & navegador.backgroundsounds & "<br>"
  response.write "Suporta Cookies?: " & navegador.cookies & "<br>"
  response.write "Suporta VBScript?: " & navegador.vbscript & "<br>"
  response.write "Suporta Javascript?: " & navegador.javascript & "<br>"
%>