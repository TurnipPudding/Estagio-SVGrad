# Interface e Modelo para Alocação de Aulas em Salas de Aula
Este projeto foi criado para auxiliar o Serviço de Graduação na alocação das aulas nas salas do Instituto de Ciências Matemáticas e de Computação (ICMC), no campus da Universidade de São Paulo (USP) de São Carlos.

Aqui, é possível baixar os arquivos necessários e tutoriais para o uso da interface, que têm como objetivo resolver o problema de otimização inteira (MIP) de Alocação de Aulas em Salas de Aula.

## Estrutura do repositório:

- archive: Pasta com scripts e pastas antigos, que não são utilizados no sistema principal da interface.
- data: Pasta com arquivos e planilhas necessárias para uso da interface. Algumas planilhas requerem manutenção constante para manter os dados atualizados.
- docs: Pasta com as documentações e manuais para utilizar a interface.
- src: Pasta com os principais arquivos necessários para instalação e execução da interface.
- tests: Pasta com arquivos de exemplo para testar o funcionamento da interface.

## Instalação:
Baixe os arquivos na pasta data. Eles serão utilizados para criar as planilhas e documentos necessários para a interface.
Existem duas maneiras de instalar e utilizar a interface: via terminal e executáveis, e via jupyter notebook e scripts.

### 1. Via Terminal
- 1.1 Dentro da pasta src deste diretório, baixe e extraia o arquivo "Via Terminal.zip". Renomear a pasta é recomendável, mas não necessário. 
- 1.2 Abra a pasta extraída
- 1.3 Dentro da pasta configs, executa o arquivo "prepara.bat", aceitando as permissões requisitadas. Após dar permissão de administrador, a versão necessária do python será instalada.
- 1.4 Após a instalação, feche a janela do terminal que não possui título. Na janela que restar, haverá uma pergunta 'Deseja finalizar o arquivo em lotes (S/N)?'. Responda com N, e deixe as dependências serem instaladas.
- 1.5 Ao final de tudo, pressione qualquer tecla para fechar a janela.
- 1.6 Com todas as dependências instaladas, basta executar o arquivo Interface.bat. É possível criar um atalho deste arquivo na área de trabalho.

### 2. Via Jupyter Notebook
- 2.1 Antes de baixar os arquivos, instale um software como VS Code ou Jupyter Notebook para executar arquivos .ipynb.
- 2.2 Com o software instalado, baixe e extraia o arquivo "Via Jupyter Notebook.zip". Renomear a pasta é recomendável, mas não é necessário.
- 2.3 Para abrir a interface, basta abrir o software baixado, abrir o arquivo "interface_run.ipynb" utilizando o software, e executar as duas células presentes no arquivo.

## Funcionamento Geral
A interface funciona utilizando planilhas com colunas pré-definidas e colunas específicas. Arquivos criados pela interface estarão em uma pasta no mesmo diretório onde os scripts foram baixados, chamada Saídas da Interface. Dentro dela, haverão outras duas pastas: Planilhas de Dados, onde os arquivos gerados somente pela interface serão salvos; e Saídas do Modelo, onde apenas arquivos gerados pela execução do modelo serão salvos.

Se alguma função da interface não for executada, ou nada acontecer ao apertar um botão, a chance de ser um problema com os documentos sendo utilizados é maior do que ser um problema apenas da interface, então verifique bem os arquivos que estão sendo utilizados.

## Manuais
Como dito anteriormente, os manuais disponibilizados na pasta docs foram criados para auxiliar futuros desenvolvimentos da interface, assim como ensinar usuários a navegar por ela.
- Relatório para a Interface: oferece um review geral da modelagem e dos scripts utilizados pelo modelo.
- Manual de Instruções para a Interface: o principal manual do usuário, com descrições e detalhes de como usar a interface, mas não fala muito de seu funcionamento.
- Documentação da Interface: principal arquivo com a documentação do arquivo interface_final.py.
- Documentação do Modelo: principal arquivo com a documentação do arquivo Modelo Universal-Copy1.py.
- Documentação da Verificação dos Dados: principal arquivo com a documentação do arquivo verificar_dados.py.
- Documentação do JSM: principal arquivo com a documentação do arquivo jupiter sheet maker.py.
