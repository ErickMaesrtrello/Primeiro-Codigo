Attribute VB_Name = "Módulo1"
Sub primeiro()
'O comando Dim(dimension) é utilizado para declarar variavel
'a variavel nome foi tipada como string(texto)
Dim nome As String

'o comando inputbox abre uma caixa de entrada de dados
'assim o usuario digita o nome e aloca na variavel nome
nome = InputBox("Digite seu nome")
'o comando Range permite selecionar uma celula na planilha do excel.
'assim selecionamos a celula A1 e adicionamos o valor que foi digitado na caixa de entrada usando a variavel nome
Range("A1").Value = nome
End Sub
