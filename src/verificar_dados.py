# -*- coding: utf-8 -*-

"""## Import das bibliotecas"""
# Imports das bibliotecas e funções utilizadas
import pandas as pd # Leitura de dados
import sys # Implementação de saídas de erro
import traceback # Implementação de inspeção para auxiliar nas saídas de erro.


"""
## Leitura dos dados
Etapa responsável por carregar os dados das planilhas de cada departamento e das salas, além de extrair listas auxiliares para análise.
"""

# Caminho do arquivo Excel passado como argumento
file_path = sys.argv[1]

# Lê o arquivo principal com todas as planilhas (salas e disciplinas)
df_completo = pd.ExcelFile(sys.argv[1])
# Extrai os nomes das planilhas presentes no arquivo
sheet_names = df_completo.sheet_names

# Lê e mescla os dados das aulas de todos os departamentos em um único DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_names[1:])
df = pd.concat(df.values(), ignore_index=True)
# print(df)  # Para depuração

# Lê os dados das salas
salas = pd.read_excel(file_path, sheet_name=sheet_names[0])
# print(salas)  # Para depuração

# Lista da capacidade de cada sala
cap_s = salas['Lugares'].tolist()
# print(cap_s)  # Para depuração

# Lista do tamanho de cada disciplina (número de inscritos)
tam_t = df['Vagas por disciplina'].tolist()
# print(tam_t)  # Para depuração

print('\nBase de Dados lida.')


"""
## Tratamento dos Dados
Etapa responsável por padronizar horários, tratar células irregulares e garantir consistência dos dados para análise posterior.
"""

# Função para converter horário no formato 'HH:MM' para valor decimal em horas
def horario_para_decimal(horario):
    # Se houver horário no formato '20h40', converte para '20:40'
    if 'h' in horario:
        horario = horario.replace('h',':')
    # Separa horas e minutos
    horas, minutos = map(int, horario.split(':'))
    # Retorna valor decimal (ex: 20:40 -> 20.67)
    return horas + minutos / 60

# Função para processar célula no formato 'Dia - HH:MM/HH:MM'
def processar_horario(celula):
    # Verifica se a célula possui horário definido
    if isinstance(celula, str) and "-" in celula:
        # Remove espaços extras
        celula = str(celula).replace(' ', '')
        # Separa dia e horários
        dia, horarios = celula.split('-')
        inicio, fim = horarios.split('/')
        # Converte horários para decimal
        start_a = horario_para_decimal(inicio)
        end_a = horario_para_decimal(fim)
        return dia, start_a, end_a
    else:
        # Célula irregular ou vazia
        return 0, 0, 0

# Lista das colunas de horários a serem processadas
colunas_horarios = ['Horário 1', 'Horário 2', 'Horário 3', 'Horário 4']
result = []
# Processa cada coluna de horários e traduz para formato padronizado
for coluna in colunas_horarios:
    resultados = df[coluna].apply(processar_horario).to_list()
    result.extend(resultados)

# Cria DataFrame com dados padronizados de todas as aulas
A = pd.DataFrame(result, columns=['Dia', 'start_a', 'end_a'])

# Salva colunas do DataFrame em listas separadas para facilitar análise
dia_a = A['Dia'].to_list()
start_a = A['start_a'].to_list()
end_a = A['end_a'].to_list()
# print(A)
# print(len(A))
# print(start_a)

# Preenche células vazias na coluna 'Sala' com valor 0
df['Sala'] = df['Sala'].fillna('0').astype(str)

# Preenche células vazias na coluna 'Turma' com valor 1
df['Turma'] = df['Turma'].fillna(1)


"""## Dados de Entrada

Bloco responsável por criar listas e dicionários que representam os índices das turmas/disciplina, salas, laboratórios, cursos e suas relações. Essas estruturas são fundamentais para a modelagem e análise dos dados de alocação de aulas e salas."""

# Lista de índices de cada turma/disciplina (0, 1, ..., n-1)
T = range(len(df['Disciplina (código)']))
# Exemplo de uso: T[38] retorna o índice da disciplina na posição 38
# Exemplo: df.loc[T[38], 'Disciplina (código)'] retorna o código da disciplina

# Lista de índices de cada sala (0, 1, ..., m-1)
S = range(len(salas['Sala']))
# Exemplo de uso: salas.loc[S[2], 'Sala'] retorna o nome da sala na posição 2

# Lista binária indicando se cada sala é laboratório (1 = sim, 0 = não)
sigma_s = [1 if salas.loc[s, 'Laboratório'] == 'Sim' else 0 for s in S]
# sigma_s[s] = 1 se sala s é laboratório

# Lista binária indicando se cada turma/disciplina precisa de laboratório (1 = sim, 0 = não)
tal_t = [0 if df.loc[t, 'Utilizará laboratório? (sim ou não)'] == 'Não' else 1 for t in T]
# tal_t[t] = 1 se turma t precisa de laboratório

# Salva os comprimentos das listas principais para evitar recomputação
lenT = len(T)   # Número de turmas/disciplina
lenA = len(A)   # Número total de aulas
lenS = len(S)   # Número de salas

# Lista dos cursos/currículos do ICMC
curriculos = ['BMACC', 'BMA', 'LMA', 'MAT-NG', 'BECD', 'BCC', 'BSI', 'BCDados']
lenC = len(curriculos)  # Número de cursos

# Dicionário Y_tc: (turma, curso) -> 1 se a turma é ministrada para o curso, 0 caso contrário
Y_tc = {(t, c): 0 for t in range(lenT) for c in range(lenC)}
for t in range(lenT):
    celula = df.loc[t, 'Curso(s)']
    celula = str(celula).replace(' ', '')  # Remove espaços extras
    # Se há mais de um curso, separa por vírgula e marca todos os cursos presentes
    if ',' in celula:
        for c in celula.split(','):
            if c in curriculos:
                Y_tc[(t, curriculos.index(c))] = 1
            
    else:
        # Se há apenas um curso, marca se estiver na lista
        if celula in curriculos:
            Y_tc[(t, curriculos.index(celula))] = 1
        
# Y_tc[(t, c)] = 1 se turma t é ministrada para curso c


"""## Dados de Preprocessamento

Bloco responsável por criar listas e dicionários auxiliares para análise e restrições do problema de alocação. Aqui são definidos agrupamentos de aulas, restrições de capacidade, conflitos de horário, uso de espaço e distâncias entre salas."""

# A_t: lista de listas, cada sublista contém os índices das aulas de uma turma/disciplina
# Exemplo: A_t[0] = [0, 1, 2] significa que as aulas 0, 1 e 2 pertencem à turma 0
A_t = []
for t in range(lenT):
    # Para cada turma, calcula os índices das aulas associadas
    A_t.append([t + (i * lenT) for i in range(int(lenA/lenT))])
# A_t é útil para mapear aulas por turma

# A_tt: igual ao A_t, mas exclui aulas sem horário definido (start_a[a] == 0)
A_tt = [[a for a in A_t[t] if start_a[a] != 0] for t in range(lenT)]
# A_tt permite filtrar apenas aulas válidas para análise

# A_s: lista de listas, cada sublista contém índices de aulas seguidas de uma mesma turma no mesmo dia
# Usada para garantir que aulas consecutivas fiquem na mesma sala
A_s = []
for t in A_t:
    # Verifica se as duas primeiras aulas da turma são no mesmo dia e estão definidas
    if dia_a[t[0]] == dia_a[t[1]] and dia_a[t[0]] != 0:
        # Verifica se o intervalo entre as aulas é compatível (não há sobreposição indevida)
        if end_a[t[0]] + 2 + 10/60 + 20/60 >= start_a[t[1]]:
            seguido = [x for x in t if start_a[x] != 0]
            if len(seguido) > 1:
                A_s.append(seguido)
# A_s é usada para restrições de alocação de aulas consecutivas

# A_c: lista de listas, cada sublista contém índices das aulas ministradas para cada curso
A_c = []
for c in range(lenC):
    # Para cada curso, inclui todas as aulas das turmas que o ministram
    A_c.append([t + (i * lenT) for i in range(int(lenA/lenT)) for t in range(lenT) if Y_tc[t, c] == 1])
# A_c[c] contém todas as aulas do curso c

# eta_as: dicionário (aula, sala) -> 1 se sala comporta a aula, 0 caso contrário
# Usa a capacidade da sala e o tamanho da turma
eta_as = {(a, s): 1 if tam_t[int((a % lenT))] <= cap_s[s] else 0 for a in range(lenA) for s in range(lenS)}
# eta_as[(a, s)] = 1 se sala s comporta aula a

# theta_aal: dicionário (aula, aula) -> 1 se há conflito de horário entre as aulas, 0 caso contrário
# Conflito ocorre se as aulas são no mesmo dia e os horários se sobrepõem
theta_aal = {(a, al): 1 if (dia_a[a] == dia_a[al] and (start_a[a] < end_a[al] and start_a[al] < end_a[a])) else 0 for a in range(lenA) for al in range(lenA)}
# theta_aal[(a, al)] = 1 se há conflito de horário entre aula a e aula al

# uso_as: dicionário (aula, sala) -> percentual de espaço vazio na sala ao alocar a aula
# Quanto menor o valor, melhor o aproveitamento da sala
uso_as = {(a, s): 100 * (1 - (tam_t[int((a % lenT))]/cap_s[s])) for a in range(lenA) for s in range(lenS)}
# uso_as[(a, s)] = percentual de espaço vazio

# dis: dicionário (sala, sala) -> distância arbitrária entre duas salas
# Usado para restrições de deslocamento entre salas
dis = {(s, sl): salas.loc[s, sl] for s in range(len(salas)) for sl in salas.columns[3:-1]}
# dis[(s, sl)] = distância entre sala s e sala sl

# Lista de salas preferencialmente vazias
pref = salas['Preferencialmente Vazia'].tolist()

"""## Dados para fixar os laboratórios e aulas com sala definida, como LEM

Bloco responsável por definir restrições específicas de alocação: aulas que exigem laboratório, aulas com sala fixa, e proibições de horários/salas. Essas estruturas garantem que certas aulas sejam obrigatoriamente alocadas em laboratórios ou salas específicas, e que restrições de uso sejam respeitadas pelo modelo.
"""

# labs: lista de índices das turmas/disciplina que precisam de laboratório (tal_t = 1)
labs = [t for t in T if tal_t[t] == 1]
# salas_labs: lista de índices das salas que são laboratórios
salas_labs = [i for i in range(len(salas['Laboratório'])) if salas.loc[i, 'Laboratório'] == "Sim"]

# ind_labs: lista de listas, cada sublista indica quais aulas de cada turma/disciplina devem ser em laboratório
# Exemplo: ind_labs[0] = [1,2] significa que as aulas 1 e 2 da turma labs[0] devem ser em laboratório
ind_labs = []
for l in labs:
    # Extrai os índices das aulas que devem ser em laboratório, removendo o texto 'Sim' e convertendo para inteiro
    valores = (df.loc[l, 'Utilizará laboratório? (sim ou não)'].replace(' ', '')).split(',')
    if "Sim" in valores:
        valores.remove("Sim")
    elif "sim" in valores:
        valores.remove("sim")
    ind_labs.append([int(item) for item in valores])

# aula_labs: lista de índices de todas as aulas que devem ser ministradas em laboratório
# Calcula o índice global da aula usando a relação entre turma e posição da aula
aula_labs = [(labs[t] + (lenT * (i-1))) for t in range(len(labs)) for i in ind_labs[t]]

# lab_tal: lista binária, cada posição indica se a aula é de laboratório (1) ou não (0)
lab_tal = [0 for _ in range(lenA)]
for a in aula_labs:
    lab_tal[a] = 1

# sala_fixa: lista de nomes de salas fixadas para cada aula; '0' indica sem sala fixa
# Se a célula de sala tem mais de um valor, seleciona o correto conforme o horário
sala_fixa = []
for a in range(lenA):
    sala_valor = str((df.loc[a % lenT, 'Sala']))
    if ', ' in sala_valor:
        if not pd.isna(df.loc[int(a / lenT), 'Horário ' + str(int(a / lenT) + 1)]):
            if len(sala_valor.split(', ')) >= (int(a / lenT) + 1):
                sala_fixa.append(sala_valor.split(', ')[int(a / lenT)])
            else:
                sala_fixa.append('0')
    elif not pd.isna(df.loc[int(a / lenT), 'Horário ' + str(int(a / lenT) + 1)]):
        sala_fixa.append(sala_valor)
    else:
        sala_fixa.append('0')

# Aplica restrição de sala fixa: zera todas as possibilidades de sala para a aula, exceto a sala fixada
for aula in range(lenA):
    if sala_fixa[aula] != '0':
        try:
            fixada = salas['Sala'].tolist().index(sala_fixa[aula])
        except ValueError:
            print(
                f"Erro de Digitação (Sala Fixada): A sala '{sala_fixa[aula]}' fixada para a disciplina {df.loc[int(aula % lenT), 'Disciplina (código)']} não existe na aba de Salas.")
            sys.exit(1)
        # 1. Salva se a turma originalmente cabia na sala fixada (0 ou 1)
        cabia_originalmente = eta_as[aula, fixada]

        # 2. Zera todas as opções para forçar o modelo a ignorar as outras salas
        for sala in range(lenS):
            eta_as[aula, sala] = 0

        # 3. Restaura a verdade física para a sala fixada
        eta_as[aula, fixada] = cabia_originalmente

# Aplica restrição de laboratório: aulas de laboratório só podem ser alocadas em salas de laboratório, e vice-versa
for aula in range(lenA):
    if lab_tal[aula] == 1:
        for sala in range(lenS):
            if sala not in salas_labs:
                eta_as[aula, sala] = 0
    else:
        for sala in range(lenS):
            if sala in salas_labs:
                eta_as[aula, sala] = 0

# sala_proibida: dicionário que armazena restrições de salas proibidas para cada aula
# Se a célula de proibição tem mais de um valor, aplica restrição para todas as salas listadas
sala_proibida = {}
for a in range(lenA):
    cell = df.loc[int(a % lenT), 'Proibir Horário ' + str(int(a / lenT) + 1)]
    if not pd.isna(cell):
        cell = str(cell)
        if ',' in cell:
            salas_proibidas = cell.split(', ')
            sala_proibida[a] = salas_proibidas
            try:
                s = salas[salas['Sala'] == sala].index[0]
                eta_as[a, s] = 0
            except IndexError:
                print(
                    f"Erro de Digitação (Sala Proibida): A sala '{sala}' proibida para a disciplina {df.loc[int(a % lenT), 'Disciplina (código)']} não existe na aba de Salas.")
                sys.exit(1)
        else:
            sala_proibida[a] = cell
            try:
                s = salas[salas['Sala'] == cell].index[0]
                eta_as[a, s] = 0
            except IndexError:
                print(
                    f"Erro de Digitação (Sala Proibida): A sala '{cell}' proibida para a disciplina {df.loc[int(a % lenT), 'Disciplina (código)']} não existe na aba de Salas.")
                sys.exit(1)

# seguidas: lista que indica se a turma tem aulas seguidas (2) ou não (1); usada para penalizar trocas de sala
seguidas = [1 for _ in range(lenT)]
for t in A_s:
    seguidas[int(t[0] % lenT)] = 2

"""## Verificação dos Dados

### Funções Auxiliares
"""


def emparelhamento_perfeito(aulas_conflito, salas_permitidas):
    # Dicionário para guardar qual aula ficou em qual sala
    sala_ocupada_por = {s: -1 for s in salas_permitidas}

    # Função recursiva que tenta achar uma sala para a aula 'u'
    def tentar_alocar(u, visitados):
        for s in salas_permitidas:
            # Se a sala comporta a aula (eta_as == 1) e ainda não tentamos ela nesta rodada
            if eta_as[u, s] == 1 and not visitados[s]:
                visitados[s] = True

                # Se a sala está livre, OU a aula que estava nela consegue ir para outra sala
                if sala_ocupada_por[s] == -1 or tentar_alocar(sala_ocupada_por[s], visitados):
                    sala_ocupada_por[s] = u
                    return True
        return False

    aulas_alocadas = 0
    # Tenta alocar cada aula do grupo de conflito
    for aula in aulas_conflito:
        visitados = {s: False for s in salas_permitidas}
        if tentar_alocar(aula, visitados):
            aulas_alocadas += 1

    # Retorna True se conseguiu alocar todas as aulas do grupo, False se faltou espaço
    return aulas_alocadas == len(aulas_conflito)


def verificar_horarios_de_conflito(grupos_de_conflitos):
    # Salas normais são todas aquelas que não são laboratórios
    salas_normais = [s for s in range(lenS) if s not in salas_labs]

    for grupo in grupos_de_conflitos:
        sucesso = emparelhamento_perfeito(grupo, salas_normais)

        if not sucesso:
            print("Há aulas com conflito de horário disputando as mesmas salas neste grupo (Falta de espaço físico).")
            print("Uma troca de horários pode ser necessária, ou a diminuição do número de vagas.")
            print(
                "Verifique se as disciplinas não foram proibidas demais, fixadas incorretamente, ou se a capacidade das salas estourou.")
            print("Em particular, o grupo de aulas e disciplinas causando problema é este:")
            for aula in grupo:
                print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")

            custom_exit()


def verificar_horarios_de_conflito_lab(grupos_de_conflitos):
    # Mapeia os índices das salas problemáticas
    nomes_salas = salas['Sala'].tolist()

    def pega_indice(nome):
        return nomes_salas.index(nome) if nome in nomes_salas else -1

    i303 = pega_indice('6-303')
    i304 = pega_indice('6-304')
    i303_304 = pega_indice('6-303/6-304')

    i305 = pega_indice('6-305')
    i306 = pega_indice('6-306')
    i305_306 = pega_indice('6-305/6-306')

    # Cria a base de laboratórios removendo essas salas específicas
    salas_base = [s for s in salas_labs if s not in [i303, i304, i303_304, i305, i306, i305_306]]

    # Opções do Bloco 1 (6-303 e 6-304)
    opcoes_bloco1 = []
    if i303 != -1 and i304 != -1:
        opcoes_bloco1.append([i303, i304]) # Configuração separada
    if i303_304 != -1:
        opcoes_bloco1.append([i303_304])   # Configuração conjunta
    if not opcoes_bloco1:
        opcoes_bloco1.append([])

    # Opções do Bloco 2 (6-305 e 6-306)
    opcoes_bloco2 = []
    if i305 != -1 and i306 != -1:
        opcoes_bloco2.append([i305, i306]) # Configuração separada
    if i305_306 != -1:
        opcoes_bloco2.append([i305_306])   # Configuração conjunta
    if not opcoes_bloco2:
        opcoes_bloco2.append([])

    # Monta as combinações físicas possíveis (4 configurações no total)
    configuracoes_fisicas = []
    for op1 in opcoes_bloco1:
        for op2 in opcoes_bloco2:
            configuracoes_fisicas.append(salas_base + op1 + op2)

    for grupo in grupos_de_conflitos:
        sucesso_geral = False
        
        # Testa o grupo contra as 4 realidades físicas possíveis
        for config in configuracoes_fisicas:
            if emparelhamento_perfeito(grupo, config):
                sucesso_geral = True
                break # Se coube em uma das configurações, é factível fisicamente!

        if not sucesso_geral:
            print("Há aulas de laboratório com conflito de horário disputando as mesmas salas neste grupo (Falta de espaço físico).")
            print("O algoritmo considerou inclusive a impossibilidade de usar salas conjuntas e individuais simultaneamente.")
            print("Uma troca de horários pode ser necessária, ou a diminuição do número de vagas.")
            print("Verifique proibições, fixações ou se a capacidade estourou.")
            print("Em particular, o grupo de aulas e disciplinas causando problema é este:")
            for aula in grupo:
                print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")

            custom_exit()


def custom_exit():
    # Obter o último frame da stack trace
    stack = traceback.extract_stack()[-2]  # O -2 pega a linha onde custom_exit foi chamada
    line_number = stack.lineno  # Pega o número da linha
    print(f"System exit da linha {line_number}")
    sys.exit()


# Aqui, trabalharemos com a ideia de uma rede de grafos. Neste caso, cada nó representa uma aula, e dois nós possuem uma aresta/conexão
# se as aulas possuem conflito de horário. Dessa forma, os nós ficam separados em clusters, mas algumas aulas são capazes de conectar dois
# clusters, como aulas das 17 às 19 que conectam o cluster das aulas das 16 às 18 com os das 18 às 20:40, que estão conectados com 
# os das 19 às 20:40. Por conta dessa natureza, na hora de fazer a verificação, haveria muito mais nós do que salas disponíveis.

group = [0 for _ in range(lenA)]
intervalos = [[8, 9.9], [10, 13], [13, 16.4], [16, 18], [19, 20.9], [21, 22.9]]  # Intervalos padrão

# Para cada aula, veja quantos intervalos ela intersecta
for a in range(lenA):
    for i_ini, i_fim in intervalos:
        
        if (start_a[a] >= i_ini or start_a[a] == 18) and end_a[a] <= i_fim and (start_a[a] > 0):
            
            group[a] += 1
        

# Se uma aula estiver em mais de um intervalo, consideramos problemática
disciplinas_problematicas = []
for a in range(lenA):
    if group[a] < 1 and start_a[a] > 0:
        if a not in disciplinas_problematicas:
            disciplinas_problematicas.append(a)



# A variável grupos_de_conflitos é uma importante lista que irá conter listas de aulas que possuem conflito entre si,
# possibilitando uma análise prévia dos dados caso tenha algo de errado com eles.
grupos_de_conflitos = []

# Para cada aula 'a'
for a in range(lenA):
    # # Verifico se a aula 'a' não tem um horário problemático, isto é, se ela não conecta dois grupos de horário separados em um único.
    if a not in disciplinas_problematicas:
        # Se ela não tem um horário problemático
        # Verifico se já fiz outras listas de conflitos
        if len(grupos_de_conflitos) > 0:
            # Se há outras listas, crio um contador para verificar a existência da aula nas outras listas
            verificar = 0
            # Para cada outra lista de conflitos
            for grupo in grupos_de_conflitos:
                # Se encontro 'a' numa outra lista, não quero adicioná-la, então contabilizo ela e paro o loop
                if a in grupo:
                    verificar += 1
                    break
            # Se eu não encontrei 'a' numa outra lista, passo para a etapa de adicionar aulas com conflito na lista atual
            if verificar == 0:
                # A lista_a é uma lista de todas as aulas que possuem conflito de horário com a aula 'a'
                lista_a = []
                # Para cada 'al', se theta_aal = 1, adiciono ela na lista
                for al in range(lenA):
                    if theta_aal[a,al] == 1:
                        lista_a.append(al)
        else:
            # Se essa é minha primeira lista de conflitos, só verifico o theta_aal e adiciono as aulas com conflito
            # A lista_a é uma lista de todas as aulas que possuem conflito de horário com a aula 'a'
            lista_a = []
            for al in range(lenA):
                if theta_aal[a,al] == 1:
                    lista_a.append(al)

    # Se a lista de conflitos foi criada corretamente, deve haver ao menos uma aula nela
    if len(lista_a) > 0:
        # Após colocar a lista na lista de grupos de conflito, eu limpo a lista atual e começo a análise da aula seguinte
        grupos_de_conflitos.append(lista_a)
        lista_a = []

# print(grupos_de_conflitos)


"""### Verificação Geral"""

# Com o grupo "bruto" de aulas com conflitos, vamos analisar de forma mais detalhada cada uma delas
# A variável laboratorios_conflito é uma lista onde cada elemento é uma lista de aulas de laboratório que possuem conflito entre si
laboratorios_conflito = []
# A variável salas_de_aula é uma lista onde cada elemento é uma lista das salas disponíveis para um determinado grupo de aulas com conflito
# Ex: Se há 3 grupos em grupos_de_conflitos, então há 3 listas de salas em salas_de_aula,
# onde cada uma dessas listas equivale a um conjunto de salas que conseguem comportar as aulas do grupo de mesmo índice, ou seja,
# salas_de_aula[0] é a lista de salas que adequam as aulas presentes em grupos_de_conflitos[0]
salas_de_aula = []
# A variável salas_de_laboratorio seguem o mesmo raciocínio, mas apenas para as aulas de laboratório
salas_de_laboratorio = []

# Com essas novas listas criadas, passamos a analisar cada grupo em grupos_de_conflitos
for grupo in grupos_de_conflitos[:]:
    
    # A variável aula_lab é uma lista que será usada para formar a lista de laboratorios_conflito
    aula_lab = []
    # A variável salas_de_aula_conflito é uma lista que será usada para formar a lista de salas_de_aula
    salas_de_aula_conflito = []
    # A variável salas_de_laboratorio_conflito é uma lista que será usada para formar a lista de salas_de_laboratorio
    salas_de_laboratorio_conflito = []
    # Para cada aula do grupo de conflito que estamos analisando
    for aula in grupo:
        # Verifico se é uma aula de laboratório
        if lab_tal[aula] == 1:
            
            # Em caso positivo, verifico se a aula foi fixada em uma sala.
            if sala_fixa[aula] != '0':
                # Verifico se a sala fixada foi um que não seja de laboratório, isto é,
                # a aula de laboratório foi fixada em uma sala que não é de laboratório.
                
                if salas['Sala'].tolist().index(sala_fixa[aula]) not in salas_labs:
                    # Caso seja o caso, envio uma mensagem de erro e interrompo o código.
                    print(
                        f"Uma aula de laboratório da disciplina {df['Disciplina (código)'][aula % lenT]} foi fixada " \
                        f"em uma sala que não é de laboratório."
                    )
                    custom_exit()
                
                elif eta_as[aula, salas['Sala'].tolist().index(sala_fixa[aula])] != 1:
                    print(
                        f"As aulas de laboratório da disciplina {df['Disciplina (código)'][aula % lenT]} foram fixadas " \
                        f"em uma sala de laboratório onde elas não cabem."
                        f"\nAlternativamente, a sala fixada também foi proíbida de ser usada por essa aula."
                    )
                    custom_exit()
                # Verifico se a aula foi fixada em uma sala onde outra aula de conflito também foi fixada.
                else:
                    
                    if sala_fixa[aula] == '6-303/6-304':
                        for aula2 in grupo:
                            if aula != aula2 and sala_fixa[aula2] != '0':
                                if sala_fixa[aula] == sala_fixa[aula2] \
                                or \
                                sala_fixa[aula2] == '6-303' \
                                or \
                                sala_fixa[aula2] == '6-304':
                                
                                    print(
                                        f"Uma aula de laboratório da disciplina {df['Disciplina (código)'][aula % lenT]} foi fixada " \
                                        f"na mesma sala (ou salas conjuntas) onde uma aula de laboratório da disciplina " \
                                        f"{df['Disciplina (código)'][aula2 % lenT]} foi fixada."
                                    )
                                    custom_exit()
                    
                    elif sala_fixa[aula] == '6-305/6-306':
                        for aula2 in grupo:
                            if aula != aula2 and sala_fixa[aula2] != '0':
                                if sala_fixa[aula] == sala_fixa[aula2] \
                                or \
                                sala_fixa[aula2] == '6-305' \
                                or \
                                sala_fixa[aula2] == '6-306':
                                
                                    print(
                                        f"Uma aula de laboratório da disciplina {df['Disciplina (código)'][aula % lenT]} foi fixada " \
                                        f"na mesma sala (ou salas conjuntas) onde uma aula de laboratório da disciplina " \
                                        f"{df['Disciplina (código)'][aula2 % lenT]} foi fixada."
                                    )
                                    custom_exit()
                    else:
                        for aula2 in grupo:
                            if aula != aula2 and sala_fixa[aula2] == 1:
                                
                                if sala_fixa[aula] == sala_fixa[aula2]:
                                    print(
                                        f"Uma aula de laboratório da disciplina {df['Disciplina (código)'][aula % lenT]} foi fixada " \
                                        f"na mesma sala (ou salas conjuntas) onde uma aula de laboratório da disciplina " \
                                        f"{df['Disciplina (código)'][aula2 % lenT]} foi fixada."
                                    )
                                    custom_exit()

            
            # Se a aula não foi fixada, eu a coloco na lista de aulas de laboratório com conflito e a retiro do grupo de conflito,
            # já que as demais aulas não precisam disputar sala com aulas de laboratório, e vice-versa
            # aula_lab.append(grupo.pop(grupo.index(aula)))
            aula_lab.append(aula)
            
            # A variável auxiliar 'aux' é uma lista que serve para guardar
            # as salas de laboratórios capazes de acomodar a 'aula' sendo analisada.
            aux = [s for s in salas_labs if eta_as[aula,s] == 1]
            # Se existe ao menos um laboratório na lista capaz de atender a aula, eu salvo a lista.
            if len(aux) > 0:
                # O uso do extend ao invés do uso de append é para evitar mais linhas de iteração para cada elemento da lista auxiliar.
                salas_de_laboratorio_conflito.extend(aux)
            else:
                # Caso contrário, envio uma mensagem de erro e interrompo o código.
                print(
                    f"Alguma(s) aula(s) de laboratório da disciplina {df['Disciplina (código)'][aula % lenT]} " \
                    f"não pode(m) ser(em) alocada(s) por conta do número de alunos."
                    f"\nAlternativamente, a sala fixada também foi proíbida de ser usada por essa aula."
                )
                
                custom_exit()
        # Para o caso de a aula não ser uma aula de laboratório.
        else:
            # Verifico se a aula está fixada em alguma sala.
            if sala_fixa[aula] != '0':
                # Para o caso da aula ter sido alocada em uma sala de laboratório.
                
                if sala_fixa[aula] in salas['Sala'].tolist() and salas['Sala'].tolist().index(sala_fixa[aula]) in salas_labs:
                    # Envio uma mensagem de erro e interrompo o código.
                    
                    print(
                        f"A aula da disciplina {df['Disciplina (código)'][aula % lenT]}" \
                        f" não é de laboratório, mas foi alocada em uma sala de laboratório."
                    )

                    custom_exit()
                
                elif eta_as[aula, salas['Sala'].tolist().index(sala_fixa[aula])] != 1:
                    print(
                        f"As aulas da disciplina {df['Disciplina (código)'][aula % lenT]} foram fixadas " \
                        f"em uma sala onde elas não cabem."
                        f"\nAlternativamente, a sala fixada também foi proíbida de ser usada por essa aula."
                    )
                    custom_exit()
                # Verifico se a aula foi fixada em uma sala onde outra aula de conflito também foi fixada.
                else:
                    # Verifico as demais aulas do grupo de conflito atual.
                    for aula2 in grupo:
                        # Se houver uma outra aula (aula2) deste grupo que está fixada na mesma sala da aula atual.
                        # if aula != aula2 and theta_aal[aula, aula2] == 1 and \
                        if aula != aula2 and \
                        sala_fixa[aula2] != '0' and \
                        salas['Sala'].tolist().index(sala_fixa[aula]) == salas['Sala'].tolist().index(sala_fixa[aula2]):
                        
                        
                            # Envio uma mensagem de erro e interrompo o código.
                            print(
                                # f"aula {aula} e aula2 {aula2}\n" \
                                f"Uma aula da disciplina {df['Disciplina (código)'][aula % lenT]} foi fixada " \
                                f"na mesma sala onde uma aula da disciplina {df['Disciplina (código)'][aula2 % lenT]} foi fixada."
                            )
                            custom_exit()

            # Utilizo novamente uma lista auxiliar com as salas que são capazes de atender àquela aula, não contando os laboratórios
            aux = [s for s in range(lenS) if eta_as[aula,s] == 1 and s not in salas_labs]
            # Se existe ao menos uma sala na lista capaz de atender a aula, eu salvo a lista
            if len(aux) > 0:
                # O uso do extend ao invés do uso de append é para evitar mais linhas de iteração para cada elemento da lista auxiliar
                salas_de_aula_conflito.extend(aux)
            else:
                # Caso contrário, envio uma mensagem de erro e interrompo o código.
                print(
                    f"Alguma(s) aula(s) da disciplina {df['Disciplina (código)'][aula % lenT]} "
                    f"não pode(m) ser(em) alocada(s) por conta do número de alunos."
                    f"\nAlternativamente, a sala fixada também foi proíbida de ser usada por essa aula."
                )
                
                custom_exit()
    for aula in grupo[:]:
        if lab_tal[aula] == 1:
            grupo.remove(aula)

    if len(grupo) == 0:
        grupos_de_conflitos.remove(grupo)
    # Nesta etapa, eu verifico quais listas foram construídas com as aulas do grupo sendo analisado.
    # Como na etapa anterior o grupo é atualizado para ficar sem as aulas de laboratório
    # não há a necessidade de alterá-lo aqui.
    # Note também que eu não checo o tamanho de 'salas_de_laboratorio_conflito', pois, se o código chegou até aqui,
    # é garantido pelas condições da etapa anterior que existam laboratórios capazes de comportar as aulas do grupo.
    # Se há elementos tanto em aula_lab quanto salas_de_aula_conflito, isto é, se no grupo haviam aulas normais e aulas de laboratório.
    if len(aula_lab) > 0 and len(salas_de_aula_conflito) > 0:
        # Adiciono a lista de aulas de laboratório na lista adequada.
        laboratorios_conflito.append(aula_lab)
        # Também adiciono as salas de laboratório que podem ser utilizadas para essas aulas.
        # Note que usar o método list(dict.fromkeys(x)) faz com que valores repetidos sejam unificados em um,
        # pois transformar uma lista em um dicionário unifica seus elementos, que depois é convertido novamente em uma lista.

        salas_de_laboratorio.append(list(dict.fromkeys(salas_de_laboratorio_conflito)))
        # O mesmo é feito para as salas de aula normais.
        salas_de_aula.append(list(dict.fromkeys(salas_de_aula_conflito)))

    # Caso haver apenas aulas comuns.
    elif len(salas_de_aula_conflito) > 0 and len(aula_lab) == 0:
        # Apenas adiciono as salas de aula que suportam o grupo.
        salas_de_aula.append(list(dict.fromkeys(salas_de_aula_conflito)))

    # De forma semelhante, caso houvesse apenas laboratórios no grupo de conflitos.
    elif len(salas_de_aula_conflito) == 0 and len(aula_lab) > 0:
        # Adiciono as aulas de laboratório conflitantes.
        laboratorios_conflito.append(aula_lab)
        # Adiciono as salas de laboratório que suportam o grupo.
        salas_de_laboratorio.append(list(dict.fromkeys(salas_de_laboratorio_conflito)))

    # Se nenhum dos casos anteriores é verdade, envio uma mensagem de erro, o grupo de turmas/disciplinas conflitantes, e interrompo o código.
    else:
        print("Alguma(s) das aulas das seguintes disciplinas não podem ser alocadas por conta do número de alunos:")
        for i in grupo:
            print(df['Disciplina (código)'][i])
        
        custom_exit()

# print("Aulas com conflito:", grupos_de_conflitos,"\n")
# print(sum([len(g) for g in grupos_de_conflitos]))
# print("Laboratórios com conflito:", laboratorios_conflito,"\n")
# print(sum([len(g) for g in laboratorios_conflito]))

# print("Salas de aula de cada grupo com conflito", salas_de_aula,"\n")
# print("Laboratórios de cada grupo com conflito", salas_de_laboratorio,"\n")


"""### Verificação por Categorias"""

# Verificação de horários conflitantes para as aulas normais.
verificar_horarios_de_conflito(grupos_de_conflitos)

# Verificação de horários conflitantes para as aulas de laboratório.
verificar_horarios_de_conflito_lab(laboratorios_conflito)

# Se nenhuma das situações de erro foi encontrada, aviso que não há nenhum erro aparente com os dados fornecidos.
print("Aparentemente, não há nenhum conflito de horário, e há salas disponíveis em todos os horários necessários.")
