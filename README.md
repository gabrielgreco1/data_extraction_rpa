# Projeto 1: Rotina de Extração de Dados
## Descrição
Este projeto consiste em uma rotina automatizada de extração de dados de um banco de dados Oracle. A rotina é executada em uma máquina virtual hospedada na AWS. Ela se conecta ao banco de dados Oracle usando as informações de conexão fornecidas, executa uma consulta SQL especificada, processa os resultados e armazena-os em um arquivo Excel. A execução da rotina pode ser agendada para ocorrer em intervalos regulares, como diariamente.

## Tecnologias Utilizadas
Python: Linguagem de programação utilizada para desenvolver a rotina.\n
oracledb: Módulo utilizado para a conexão com o banco de dados Oracle.
pandas: Biblioteca usada para a manipulação de dados em formato tabular.
openpyxl: Biblioteca utilizada para a manipulação de arquivos Excel.
datetime: Módulo para manipulação de datas e horários.
AWS: Plataforma de computação em nuvem utilizada para hospedar a máquina virtual onde a rotina é executada.

## Utilização
Imagine que o Departamento de Recursos Humanos (RH) da empresa precisa de relatórios diários sobre a presença e ausência dos funcionários, ou informações sobre folha de pagamento e benefícios (que é o caso deste projeto). A rotina automatizada pode ser configurada para executar todos os dias de manhã cedo, ou em qualquer outro período escolhido pelos solicitantes, pois como roda em uma máquina virtual podemos programar as tasks para qualquer momento. Os resultados são extraídos do banco de dados Oracle e armazenados em um arquivo Excel formatado. Esse arquivo é, então, automaticamente salvo em uma pasta compartilhada na AWS.

# Projeto 2: Interface Gráfica para Extração de Dados
## Descrição
Este projeto envolve uma aplicação de interface gráfica hospedada em uma máquina virtual na AWS. A aplicação permite que os usuários de diversos setores da empresa selecionem o tipo de relatório que desejam extrair, insiram a data do período desejado e, em seguida, executem a extração dos dados. Os resultados são armazenados em arquivos Excel separados, com nomes baseados no tipo de relatório e na data do período.

## Tecnologias Utilizadas
tkinter: Biblioteca padrão do Python para criação de interfaces gráficas.
...

## Utilização 
A aplicação de interface gráfica é especialmente útil para os membros de todos os setores de uma empresa, pois simplifica o processo de extração de relatórios. Eles podem usar a aplicação para selecionar o tipo de relatório (por exemplo, "Benefícios" ou qualquer outra tabela do banco de dados) e inserir a data do período que desejam analisar. Após a execução, os dados são extraídos do banco de dados Oracle, processados e armazenados em arquivos Excel. Esses arquivos são automaticamente salvos em uma pasta compartilhada na AWS, onde os membros do setor podem acessá-los facilmente.

Conclusão
Esses dois projetos automatizados fornecem soluções eficazes para extração automática de informações do banco de dados da sua empresa. Eles permitem a extração e armazenamento de relatórios diários de forma eficiente e conveniente. Com a hospedagem na AWS, esses projetos podem ser executados de forma confiável em uma máquina virtual, garantindo a disponibilidade dos relatórios quando necessário. Certifique-se de configurar corretamente as informações de conexão, consultas SQL e outros parâmetros para atender às necessidades específicas do seu ambiente na AWS.
