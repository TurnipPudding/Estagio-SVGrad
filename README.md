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
Dentro da pasta src deste diretório, baixe todos os arquivos disponíveis, com exceção de "interface_run.ipynb", incluindo a pasta "configs".
Coloque os arquivos baixados em uma pasta, junto com a pasta "configs". Dentro da pasta configs, executa o arquivo "prepara.bat", e aceite as permissões requisitadas. Após dar permissão de administrador, a versão necessária do python será instalada. Após a instalação, feche a janela do terminal que não possui título. Na janela que restar, haverá uma pergunta 'Deseja finalizar o arquivo em lotes (S/N)?'. Responda com N, e deixe as dependências serem instaladas. Ao final de tudo, pressione qualquer tecla para fechar a janela. Com todas as dependências instaladas, basta executar o arquivo Interface.bat.

### 2. Via Jupyter Notebook
Antes de baixar os arquivos, instale um aplicativo como VSCode ou Jupyter Notebook para executar arquivos .ipynb. Com o aplicativo instalado, baixe os arquivos .py e .ipynb na pasta src. Não há necessidade de baixar a pasta configs ou o arquivo "Interface.bat". Com todos os arquivos salvos em uma pasta, abra o arquivo interface_run.ipynb usando o aplicativo instalado e execute as duas células.

## Funcionamento Geral
A interface funciona utilizando planilhas com colunas pré-definidas e colunas específicas. Arquivos criados pela interface estarão em uma pasta no mesmo diretório onde os scripts foram baixados, chamada Saídas da Interface. Dentro dela, haverão outras duas pastas: Planilhas de Dados, onde os arquivos gerados somente pela interface serão salvos; e Saídas do Modelo, onde apenas arquivos gerados pela execução do modelo serão salvos.

Se alguma função da interface não for executada, ou nada acontecer ao apertar um botão, a chance de ser um problema com os documentos sendo utilizados é maior do que ser um problema apenas da interface, então verifique bem os arquivos que estão sendo utilizados.
