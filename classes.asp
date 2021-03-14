<%

Class ClassName 'Abertura do bloco de classe

	' Usar prefixo m_ para indicar variáveis da classe, pois não se pode repetir o nome no método Get
	' Não se pode atribuir valores na declaração
	private m_propriedade1 ' Valor privado, não pode ser acessado externamente
	dim m_propriedade2 'Por padrão os valores são públicos e podem ser acessados externamente
	public m_propriedade3 ' Declaração explicita de variável publica

	' Classes tem 2 eventos padrão, um executado ao instanciar e outro ao finalizar um objeto (atribuindo Nothing)
	' Evento de Inicialização
	Private Sub Class_Initialize
		m_propriedade1 = 1
		m_propriedade2 = 2
		m_propriedade3 = 3
	End Sub
	' Evento de finalização - Lembrar sempre de fechar um objeto atribuindo NOTHING
	' Objetos são fechados automaticamente ao finalizar a pagina, mas liberar manualmente garante a memoria mais cedo
	Private Sub Class_Terminate
		response.write "...Encerando Instacia..."
	End Sub

	'---------------------------------------------------------------------

	' Você pode declarar Property Get, Let(Equivalente ao Set em outras linguagens) e Set(Especifico para objetos)
	Public Property Get propriedade1
		' Retorno do Get, similar a uma função
		propriedade1 = m_propriedade1
	End Property

	Public Property Let propriedade1(input)
		' Modifica a propriedade interna de acordo com o input
		m_propriedade1 = input
	End Property

	' Mesma notação do Let, porém usando Set na atribuição da propriedade
	Public Property Set propriedade3(inputObject)
		Set m_propriedade3 = inputObject
	End Property

	'---------------------------------------------------------------------

	' Métodos de função podem ser públicos ou privados, serem Subs ou Functions
	Public Sub Shazam(input)
		response.write input
	End Sub

	Public Function Shazam2(input)
		Shazam2 = "Seu input foi: " & input
	End Function

End Class ' Fechamento do bloco de classe

Dim object
Set object = new ClassName ' Instancia da Classe

object.propriedade1 = "teste"      'Alterando propriedade privada por Let
response.write object.propriedade1 'Acessando propriedade privada por Get
response.write "<br><br>"

object.m_propriedade2 = "Não é mais 2"
response.write object.m_propriedade2 ' Acessando propriedade publica
response.write "<br><br>"

' Atribuindo um objeto via Property Set
Set object.propriedade3 = Server.CreateObject("Scripting.FileSystemObject")

Set object = Nothing 'Fechando instancia explicitamente (CORRETO)
Set object = new ClassName ' Abrindo Instancia sem fechar
' Código vai gerar 2 eventos de saida, visto que a segunda instância vai se fechar sozinha

%>