<%

Class ClassName 'Abertura do bloco de classe

	' Usar prefixo m_ para indicar vari�veis da classe, pois n�o se pode repetir o nome no m�todo Get
	' N�o se pode atribuir valores na declara��o
	private m_propriedade1 ' Valor privado, n�o pode ser acessado externamente
	dim m_propriedade2 'Por padr�o os valores s�o p�blicos e podem ser acessados externamente
	public m_propriedade3 ' Declara��o explicita de vari�vel publica

	' Classes tem 2 eventos padr�o, um executado ao instanciar e outro ao finalizar um objeto (atribuindo Nothing)
	' Evento de Inicializa��o
	Private Sub Class_Initialize
		m_propriedade1 = 1
		m_propriedade2 = 2
		m_propriedade3 = 3
	End Sub
	' Evento de finaliza��o - Lembrar sempre de fechar um objeto atribuindo NOTHING
	' Objetos s�o fechados automaticamente ao finalizar a pagina, mas liberar manualmente garante a memoria mais cedo
	Private Sub Class_Terminate
		response.write "...Encerando Instacia..."
	End Sub

	'---------------------------------------------------------------------

	' Voc� pode declarar Property Get, Let(Equivalente ao Set em outras linguagens) e Set(Especifico para objetos)
	Public Property Get propriedade1
		' Retorno do Get, similar a uma fun��o
		propriedade1 = m_propriedade1
	End Property

	Public Property Let propriedade1(input)
		' Modifica a propriedade interna de acordo com o input
		m_propriedade1 = input
	End Property

	' Mesma nota��o do Let, por�m usando Set na atribui��o da propriedade
	Public Property Set propriedade3(inputObject)
		Set m_propriedade3 = inputObject
	End Property

	'---------------------------------------------------------------------

	' M�todos de fun��o podem ser p�blicos ou privados, serem Subs ou Functions
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

object.m_propriedade2 = "N�o � mais 2"
response.write object.m_propriedade2 ' Acessando propriedade publica
response.write "<br><br>"

' Atribuindo um objeto via Property Set
Set object.propriedade3 = Server.CreateObject("Scripting.FileSystemObject")

Set object = Nothing 'Fechando instancia explicitamente (CORRETO)
Set object = new ClassName ' Abrindo Instancia sem fechar
' C�digo vai gerar 2 eventos de saida, visto que a segunda inst�ncia vai se fechar sozinha

%>