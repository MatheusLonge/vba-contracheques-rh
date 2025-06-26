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
