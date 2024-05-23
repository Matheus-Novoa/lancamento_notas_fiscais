# lancamento_notas_fiscais
Bot para lançamento de notas fiscais em sistema contábil.

O Bot automatiza o carregamento dos dados dos clientes vindos de uma planilha eletrônica e faz o processo de cadastramento das notas no sistema.

Foi utilizado o framework em python (*full code*) Botcity para o desenvolvimento. Para este projeto foram necessários conhecimentos de manipulação de dados e de recursos do próprio framework, como visão computacional através de capturas de tela para a manipulação do cursor do computador.

O Bot ainda necessita da intervenção do usuário para validar cada cadastro (através do ok em uma janela de diálogo). Em futuras versões pretendo implementar alguns testes de validação, utilizando informações da tela através de OCR para mitigar possíveis erros no decorrer do processo. Com estas implementações, o objeto é tornar o processo totalmento automatizado, ou seja, sem que seja necessária a interação do usuário com o Bot.