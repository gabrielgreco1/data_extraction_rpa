# Documentação dos Projetos de automação em Python
## Projeto 1: Automação com Extração de Dados e Manipulação de Excel (extractor.py)
O primeiro projeto consiste em uma automação que realiza a extração de dados de um banco de dados Oracle e adiciona estas informações a um arquivo Excel. Utilizando as bibliotecas oracledb para conexão e consultas ao banco Oracle, pandas para manipulação de dados, openpyxl para gerenciar os arquivos Excel e datetime para controlar datas, este código se destaca pela sua eficiência em agregar dados ao arquivo Excel de forma automatizada.

A execução desse código ocorre da seguinte forma: se for o primeiro dia do mês, o arquivo Excel é limpo e os dados são extraídos da query montada diretamente no código. Caso não seja o primeiro dia do mês, os dados são apenas adicionados (append) ao final do arquivo, mantendo os dados existentes. Isso permite a manutenção de um histórico dos dados que são atualizados diariamente, sempre agregando novas informações sem a necessidade de intervenção manual.

Este projeto, portanto, é uma solução automatizada que oferece um método simples e eficiente para a extração e manipulação de dados. A implementação direta da query no código facilita a manutenção e torna o programa robusto, uma vez que as possíveis alterações estão restritas ao código fonte, sem dependência de entradas externas.

## Projeto 2: Automação com Interface Gráfica, Consultas Parametrizadas e Seleção de Tabela (extractor_gui.py)
O segundo projeto leva a automação um passo adiante, ao incorporar uma interface gráfica construída com a biblioteca TKinter. Esta interface permite ao usuário não apenas inserir uma data que será usada para filtrar a consulta ao banco de dados Oracle, mas também selecionar a tabela de onde deseja extrair os dados.

O funcionamento do projeto é simples e intuitivo. Por meio da interface gráfica, o usuário tem a opção de selecionar de qual tabela deseja extrair as informações. Após essa seleção, é permitido ao usuário fornecer uma data, que é utilizada para filtrar os dados na consulta ao banco de dados. Este processo proporciona uma maior flexibilidade e personalização da extração de dados, uma vez que o usuário pode escolher de qual tabela e qual período de tempo deseja obter os dados.

As informações resultantes da consulta ao banco de dados, então, são adicionadas a um arquivo Excel. Este processo é similar ao do primeiro projeto, com a diferença que aqui, a consulta é personalizada com base na seleção do usuário. Isso permite uma interação mais dinâmica e eficiente entre o usuário e a aplicação, tornando o processo de extração de dados mais flexível e adequado às necessidades específicas do usuário.

Assim, este segundo projeto, utilizando as bibliotecas oracledb, pandas, openpyxl, datetime e TKinter, oferece uma solução robusta e personalizável para a extração e manipulação de dados. Ele combina a eficiência da automação com a flexibilidade de uma interface de usuário amigável, permitindo que os usuários personalizem suas consultas de acordo com suas necessidades específicas.

## Projeto 3: Automação com Extração de Dados, Verificação de Duplicidade e Intervalo de Tempo (extractor_interval.py)
O terceiro projeto em Python amplia as funcionalidades dos projetos anteriores, introduzindo novas características: a verificação de dados duplicados e a extração de dados em um intervalo de tempo específico. Utilizando as bibliotecas oracledb, pandas, openpyxl e datetime, este código oferece uma solução mais completa e robusta para a extração e manipulação de dados de um banco de dados Oracle para um arquivo Excel.

O processo de extração de dados desse projeto ocorre de maneira inteligente e eficiente. Ele busca informações do banco de dados Oracle considerando um intervalo de tempo específico, que vai do primeiro dia do mês até o dia atual. Isso proporciona uma flexibilidade e eficiência únicas ao garantir que os dados coletados estejam sempre dentro do período de tempo desejado.

Após a extração dos dados, o código se encarrega de verificar a existência de duplicatas no arquivo Excel. Caso haja duplicatas, estas são excluídas, garantindo que os dados apresentados estejam sempre atualizados e sem redundâncias. Se não houver duplicatas, o código simplesmente adiciona os novos dados ao final do arquivo. Essa característica garante a integridade dos dados e evita informações redundantes ou desatualizadas.
