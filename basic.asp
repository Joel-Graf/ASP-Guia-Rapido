<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ASP - Sintaxe Básica</title>
  <style>
    .titulo {
      background-color: red;
    }
  </style>
</head>
<body>
  <h1>Guia rápido da linguagem</h1>

  <h2 class='titulo'>Sintaxe Basica</h2>
  <%
    response.write "Hello" 'Renderiza texto
    response.write("World" & "!!") 'Outra forma válida. & -> Uso de concatenação
    response.write("<br>")
    ' A forma abaixo é uma abrebviação
  %>
  <%="Hello World <br>"%>
  <%
    'A linguagem não é case-sensitive (Não existe diferença entre letra maiúscula e minúscula)
    dim variavel 'Declaração variável
    variavel = 1 'Atribuição
    A = 2 'Declaração + Atribuição

    response.write "Resultado: " & a + VarIavEl & "<br>"  'Juntando tudo que foi aprendido
  %>

  <h2 class='titulo'>Operadores<h2>
  <a
    href='http://www.tizag.com/aspTutorial/aspOperators.php'
    target='_blank'
  >Refêrencia</a>

  <h2 class='titulo'>Condicionais<h2>
  <%
    'Simple IF
    if true then
      response.write "Simple IF <br>"
    end if

    'IF/ELSE
    If false then
      response.write "IF <br>"
    else
      response.write "ELSE <br>"
    end if

    'Switch case
    select case 9
    case 1,2,3
      response.write("Foi 1 ou 2 ou 3 <br>")
    case 5
      response.write("Foi 5 <br>")
    case else
      response.write("Não foi nenhum valor <br>")
    end select
  %>

  <h2 class='titulo'>Loops</h2>
  <%
    'For loop padrão
    For i = 0 To 6 'Inclui o valor final (repete 7x)
      if i = 4 then exit for 'Sai do laço For
      response.write("For loop: " & i & "<br>")
    Next

    'For loop decrescente (Pode alterar incremento pelo step)
    For i = 8 To 0 Step -2
      response.write "For loop (-2): " & i & "<br>"
    Next

    'While - Valida condição na entrada
    counter = 1
    Do While counter < 10
      if counter = 8 then exit do 'Sai do laço Do While
      response.write "While: " & counter & "<br>"
      counter = counter + 1 'Counter++ é inválido :c
    Loop

    'Do While - Não valida na entrada
    counter = 2
    Do
      response.write "Do While: " & counter & "<br>"
      counter = counter + 1
    Loop While counter < 1

    'Until - Sai quando a condição é atingida
    i = 5
    Do until i = 10
      response.write "Until: " & i & "<br>"
      i = i + 1
    Loop

    'Do Until
    i = 5
    Do
      response.write "Do Until: " & i & "<br>"
      i = i + 1
    Loop Until i = 10
  %>

  <h2 class='titulo'>Funções e Procedures</h2>
  <%
    'Declaração de função, com parametros
    Function RetornaDiaSemana(data)
      dim dias (8)

      dias(1) = "Domingo"
      dias(2) = "Segunda"
      dias(3) = "Terça"
      dias(4) = "Quarta"
      dias(5) = "Quinta"
      dias(6) = "Sexta"
      dias(7) = "Sábado"

      diaSemana = Weekday(data)
      RetornaDiaSemana = dias(diaSemana) 'Retono da função!
    End Function
    response.write RetornaDiaSemana("10/02/2020") & "<br>"

    'Procedures, função sem retorno
    Sub EscrevaTexto(texto)
      response.write texto
    end sub
    EscrevaTexto("Oi eu sou um texto gerado em Procedure!")
    response.write "<br>"
  %>

  <h2 class='titulo'>Prioridade de Execu��o</h2>
  <!-- Tag script � executada depois do < % -->
  <script language=VBScript RUNAT=Server>
	'response.write('Antes')
  </script>
  <%="Depois"%>

  <h2 class='titulo'>Interrompendo Script</h2>
  <%
    'O script pode ser interrompido utilizando response.end
    'Isso faz com que TODO script abaixo seja interrompido, por isso está no final da página
    response.write "Linha 1"
    response.write "Linha 2"
    response.end
    response.write "Linha 3"
    response.write "Linha 4"
  %>

</body>
</html>