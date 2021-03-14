<%
  'Cria e escreve um arquivo no servidor 
  'Cria os objetos de File System
  dim fs, fname
  set fs = server.createObject("Scripting.FileSystemObject") 
  set fname = fs.createTextFile("C:\Users\joejo\Desktop\Programacao\Asp\joel\teste.txt", true) 
  fname.writeLine("Alô mundo!") 'Escreve no arquivo
  
  'Fecha os recursos
  fname.close
  set fname = Nothing
  set fs = Nothing

  response.write "Arquivo criado."
%>
<br><br>

<%
  'Verifica existência do arquivo
  set fs = server.createObject("Scripting.FileSystemObject")

  if fs.folderExists("c:\temp") = true then
    response.write "pasta c:\temp existe!"
  else
    response.write "pasta c:\temp não existe." 
  end if

  set fs = Nothing
%>
<br><br>


<%
  'Verifica existência da pasta
  set fs = server.createObject("Scripting.FileSystemObject")
  set pasta = fs.getFolder("C:\Users\joejo\Desktop\Programacao\Asp\joel\folderTest")
  'Pega a data de criação da pasta por uma propriedade do objeto
  response.write("Pasta criada em: " & pasta.DateCreated) 
  set pasta = Nothing
  set fs = Nothing
%>
<br><br>

<%
  'Verifica propiedades de disco rígido
  set fs = server.createObject("Scripting.FileSystemObject")
  set disco = fs.getDrive("C:")
  response.write("Tamanho total de c: em bytes: " & disco.totalsize)
  set disco = nothing
  set fs = nothing
%>
<br><br>

<%
  'Verifica propriedades de um arquivo
  dim fs, arquivo
  set fs = server.createObject("Scripting.FileSystemObject")
  caminho = server.mappath("basic.asp")
  set arquivo = fs.getfile(caminho)
  response.write("Este arquivo foi criado em: " & arquivo.dateCreated)
  set arquivo = nothing
  set fs = nothing
%>
<br><br>

<% 
  'Mostra caminho físico do arquivo
  caminho = Server.MapPath("basic.asp")
  response.write "O arquivo basic.asp está na pasta " & caminho
%>
<br><br>