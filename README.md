# vba-contracheques-rh
  Automação em VBA para extração de contracheques históricos em arquivos CSV<br>
  Este projeto foi desenvolvido para automatizar a busca de contracheques históricos de um funcionário, usando um script em **VBA no Excel**. A Automação percorre arquivos CSV organizados     por mês e ano (1994 a 2018), localiza o nome e data inicial informados (melhores parâmetros que considerei afim da iteração considerar como bloco e extrair as linhas relevantes). Ainda     faço uso do mesmo, com futuras atualizações.
## Objetivo
  Evitar o trabalho manual de abrir centenas de arquivos, localizar registros de funcionários, se estava em mandato eleitoral, cedido para outro órgão público entre outras exceções vulneráveis a retrabalho e erros humanos.
## Funcionalidades
  - Leitura de arquivos .txt/.csv organizados por pasta (ex: 'G:\...1994\JAN1994')
  - Busca por nome e data de início (Os melhores parâmetros para busca para reconhecimento do contracheque como um bloco e sua extração, visto que a alteração de matrícula do funcionário poderia ocorrer e nesse caso a trativa de exceção ficaria extensa, como mostrarei logo abaixo e no script)
  - Copia os blocos relevantes para uma aba Excel ('Plan10') Obs: Por necessidade específica do trabalho que exerço evito a exportação direta para outros diretórios do setor.
  - Informa arquivos/mês/ano não encontrados
  - Organiza os dados de forma clara e cronológica
## Tecnlogias utilizadas
  - **VBA (Visual Basic for Applications)**
  - **Excel**
  - **Manipulação de arquivos via 'Open For Input'**
## Aprendizados
  - Aplicação prática de estruturas de repetição e controle de fluxo em VBA
  - Manipulação de arquivos e strings
  - Automatização de tarefas administrativas reais
## Exemplos simulado de 1(um) contracheque processado pelo sistema:
![image](https://github.com/user-attachments/assets/6e8bf84b-fb2d-4680-a8f5-08f1bdb3a3b8)
Obs: As informações do exemplo citado acima, são fictícias, meramente ilustrativas.
## 🔍 Motor de busca
![image](https://github.com/user-attachments/assets/4e28e165-53d9-465c-82c5-380dc1038287)
## Resultados da busca
![image](https://github.com/user-attachments/assets/a9feae35-741e-4b75-b784-aab639f9a04f)
## 🧠 Como importar no Excel (VBA)
  1. Abra o Excel (qualquer planilha).
  2. Pressione `ALT + F11` para abrir o **Editor do VBA**
  3. No menu superior, clique em `File > Import File...`
  4. Selecione o arquivo `vba-contracheques-rh`.
  5. O código será adicionado automaticamente a um módulo do VBA
  6. Volte para o Excel, ajuste as células `A2` (nome do funcionário) e B2 (data de admissão) na aba `PLAN2`(Essa é a aba do motor de busca), e pressione `F5` no editor para rodar o código.
## ⚠️ POSSÍVEIS ERROS E OBSERVAÇÕES
  - **Arquivo não encontrado:** se o caminho `"G:\ESPELHO DA FOLHA SDE.SUDIC.CIS\..."` estiver incorreto ou inacessível, nenhum dado será lido. Certifique-se de ajustar o caminho ou montar uma estrutura similar localmente. Eu recomendo que você altere esse caminho para uma pasta sua, com um nome mais simples, utilizo essa, pois outros colegas a utilizam.
  - **Estrutura de pastas obrigatória:** os arquivos devem estar organizados em pastas por ano/mês com nomes como `JAN1994`, `FEV1994`, etc.
  - **Funcionário não localizado:** se o nome digitado em `A2` não estiver em nenhum arquivo, o script não colará nada na planilha. E     aqui está um fator do qual não pude simplificar em termos de trativa de erro, exceções. Note:<br>
```vba
If copiando Then
                        ' Verificar se encontramos o final do bloco
                        If InStr(1, linha, "MG.CONSIG.") > 0 Or InStr(1, linha, "OUTROS AFASTAMENTOS") > 0 Or InStr(1, linha, "EXERCICIO DE MANDATO ELETIV") > 0 Then
                            copiando = False
                            
                            ' Colar o bloco apenas se `blocoEncontrado` for verdadeiro
                            If blocoEncontrado Then
                                ' Colar as duas primeiras linhas como cabeçalho
                                ws.Cells(linhaAtual, 1).Value = primeiraLinha
                                linhaAtual = linhaAtual + 1
                                ws.Cells(linhaAtual, 1).Value = segundaLinha
                                linhaAtual = linhaAtual + 1
                                
                                ' Colar o bloco de linhas copiadas
                                ws.Cells(linhaAtual, 1).Value = linhasCopiadas
                                linhaAtual = linhaAtual + 1
                                
                                ' Limpar as variáveis do bloco
                                linhasCopiadas = ""
                                blocoEncontrado = False
                            End If
                        End If
```
Perceba que se o final do bloco **não contiver essas expressões** ("MG.CONSIG.", "OUTROS AFASTAMENTOS", "EXERCÍCIO DE MANDATO ELETIV"), o script pode:
  1. Parar de copiar antes da hora
  2. Ignorar blocos válidos
  3. Colar dados incompletos
  4. E o que ocorreu muitas vezes comigo: **Copiar dados do(s) servidor(es) seguintes**, aglutinando informações
Infelizmente isso obriga mesmo com mais argumentos em campos de busca, incluir mais palvras-chave, relacionada a possíveis situações do funcionário em folha de pagamento
  - **Dica:** utilize o arquivo fictício para testar em um ambiente simulado.
## Arquivo de testes (fictício)<br>
Recomendo utilizar o exemplo para testar a automação:
[ABR1997.txt](https://github.com/user-attachments/files/20931464/ABR1997.txt)
Para usar:
1. Coloque os arquivos de teste localmente, por exemplo em `C:\Testes\1994\JAN1994`
2. Altere a linha no código VBA:
```vba
caminhoBase = "C:\\Testes\\"
```
## Contato
  - LinkedIn: [Matheus Longe](https://www.linkedin.com/in/matheus-longe-aa1a221b5)
  - E-mail: m.longe.dev@gmail.com

