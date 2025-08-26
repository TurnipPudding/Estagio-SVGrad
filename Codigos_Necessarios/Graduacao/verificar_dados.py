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
for s in range(len(df['Sala'])):
    if pd.isna(df.loc[s, 'Sala']):
        df.loc[s, 'Sala'] = '0'

# Preenche células vazias na coluna 'Turma' com valor 1
for d in range(len(tam_t)):
    if pd.isna(df.loc[d, 'Turma']):
        df.loc[d, 'Turma'] = 1


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
    # Se há mais de um curso, separa por vírgula e marca todos os cursos presentes
    if ',' in celula:
        for c in celula.split(', '):
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
        fixada = salas['Sala'].tolist().index(sala_fixa[aula])
        for sala in range(lenS):
            eta_as[aula, sala] = 0
        eta_as[aula, fixada] = 1

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
            for sala in salas_proibidas:
                s = salas[salas['Sala'] == sala].index[0]
                eta_as[a, s] = 0
        else:
            sala_proibida[a] = cell
            s = salas[salas['Sala'] == cell].index[0]
            eta_as[a, s] = 0

# seguidas: lista que indica se a turma tem aulas seguidas (2) ou não (1); usada para penalizar trocas de sala
seguidas = [1 for _ in range(lenT)]
for t in A_s:
    seguidas[int(t[0] % lenT)] = 2

"""## Verificação dos Dados

### Funções Auxiliares
"""

def verificar_horarios_de_conflito(grupos_de_conflitos, salas_de_aulas):
    # Para cada grupo de conflito
    for grupo in grupos_de_conflitos:
        # Eu separo as aulas conflitantes em grupos baseados em quais salas elas cabem
        # E também separo as salas que suportam os grupos correspondentes
        grupo30 = [] # Aulas com até 30 alunos
        salas30 = []
        grupo45 = [] # Aulas com 31 a 45 alunos
        salas45 = []
        grupo73 = [] # Aulas com 46 a 73 alunos
        salas73 = []
        grupo77 = [] # Aulas com 74 a 77 alunos
        salas77 = []
        grupo80 = [] # Aulas com mais de 78 alunos
        salas80 = []

        # Vou definir algumas listas com os índices de algumas salas que comportam essas categorias.
        lista4 = salas.index[salas['Sala'].isin(['5-102', '3-102', '3-103', '3-104'])].tolist()
        lista3 = salas.index[salas['Sala'].isin(['3-009', '3-010', '3-011', '3-012', '5-002'])].tolist()
        lista2 = salas.index[salas['Sala'].isin(['5-103', '5-104'])].tolist()
        lista1 = salas.index[salas['Sala'].isin(['5-001', '5-003', '5-004', '5-101'])].tolist()
        lista0 = salas.index[salas['Sala'].isin(['4-001', '4-003', '4-005'])].tolist()

        
        # A variável verificar_salas_fixadas é uma Lista das listas de salas fixadas pelas aulas.
        # Ex: Se duas aulas foram alocadas em salas grandes, então a primeira sublista contém quais dessas salas foram alocadas.
        # Isso ajudará nos momentos em que uma aula pequena for fixada em uma sala grande, pois na hora de verificar a
        # demanda de salas grandes, o código contabiliza essa fixação, ou seja, se eu tiver três aulas que precisam de salas grandes
        # e uma aula pequena fixada em uma sala grande, mas tenho apenas três salas grandes, o código perceberá esse conflito e mandará
        # um aviso ao usuário de que alguma aula foi alocada em uma sala requisitada por outras aulas.
        verificar_salas_fixadas = [[],[],[],[],[]]

        # Para cada aula no grupo em análise
        for aula in grupo:
            # Verifico em qual categoria a aula se enquadra, adicionando-a no grupo e salvando as salas que comportam a aula.
            if tam_t[int(aula % lenT)] <= 30:
                grupo30.append(aula)
                salas30.extend([s for s in salas_de_aulas[grupos_de_conflitos.index(grupo)] if eta_as[aula,s] == 1])
                salas30 = list(dict.fromkeys(salas30))
                last_entered = 4
            elif tam_t[int(aula % lenT)] <= 45 and tam_t[int(aula % lenT)] > 30:
                grupo45.append(aula)
                salas45.extend([s for s in salas_de_aulas[grupos_de_conflitos.index(grupo)] if eta_as[aula,s] == 1])
                salas45 = list(dict.fromkeys(salas45))
                last_entered = 3
            elif tam_t[int(aula % lenT)] <= 73 and tam_t[int(aula % lenT)] > 45:
                grupo73.append(aula)
                salas73.extend([s for s in salas_de_aulas[grupos_de_conflitos.index(grupo)] if eta_as[aula,s] == 1])
                salas73 = list(dict.fromkeys(salas73))
                last_entered = 2
            elif tam_t[int(aula % lenT)] <= 77 and tam_t[int(aula % lenT)] > 73:
                grupo77.append(aula)
                salas77.extend([s for s in salas_de_aulas[grupos_de_conflitos.index(grupo)] if eta_as[aula,s] == 1])
                salas77 = list(dict.fromkeys(salas77))
                last_entered = 1
            elif tam_t[int(aula % lenT)] <= 124 and tam_t[int(aula % lenT)] > 77:
                grupo80.append(aula)
                salas80.extend([s for s in salas_de_aulas[grupos_de_conflitos.index(grupo)] if eta_as[aula,s] == 1])
                salas80 = list(dict.fromkeys(salas80))
                last_entered = 0
            else:
                # Se eu não consegui colocar aquela aula em nenhuma das salas disponíveis, novamente notifico o erro.
                print(f"Não há sala capaz de comportar a disciplina {df.loc[aula % lenT, 'Disciplina (código)']}.")
                custom_exit()
            # Se a aula possui uma sala fixada, preciso editar a lista de verificação.
            if sala_fixa[aula] != '0':
                
                if salas[salas['Sala'] == sala_fixa[aula]].index[0] in lista0:
                    # Se a sala fixada é uma das do bloco 4, aumento o valor dos elementos da lista verificar_fixadas desde o índice
                    # equivalente ao grupo80, ou seja, de [0,0,0,0,0] para [1,1,1,1,1].
                    # verificar_fixadas[0:] = [x + 1 for x in verificar_fixadas[0:]]
                    # Também adiciono o índice da sala fixada na lista de mesmo índice do último grupo que a aula foi colocada, ou seja,
                    # se uma aula que pertence ao grupo77 foi fixada em uma sala do bloco 4, eu adiciono a sala na segunda posição da
                    # lista verificar_salas_fixadas, então de [[],[],[],[],[]] foi para [[],[5],[],[],[]].
                    # verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == str(df.loc[int(aula % lenT), 'Sala'])].index[0])
                    verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == sala_fixa[aula]].index[0])
                # As demais condições são equivalentes no raciocínio.
                
                elif salas[salas['Sala'] == sala_fixa[aula]].index[0] in lista1:
                    
                    verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == sala_fixa[aula]].index[0])
                
                elif salas[salas['Sala'] == sala_fixa[aula]].index[0] in lista2:
                    
                    verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == sala_fixa[aula]].index[0])
                
                elif salas[salas['Sala'] == sala_fixa[aula]].index[0] in lista3:
                
                    verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == sala_fixa[aula]].index[0])
                
                elif salas[salas['Sala'] == sala_fixa[aula]].index[0] in lista4:
                
                    verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == sala_fixa[aula]].index[0])





        # Formados os subgrupos de aulas conflitantes separados, vou analisar a existência de aulas disponíveis para este grupo.
        # Ex: Se eu tenho 3 aulas de mesmo horário, mas apenas 2 salas disponíveis para alocá-las, há um problema com o grupo.
        verificar_aulas = [grupo80, grupo77, grupo73, grupo45, grupo30]
        verificar_salas = [salas80, salas77, salas73, salas45, salas30]
        
        # Para cada lista em verificar_salas_fixadas, isto é, para cada lista de salas fixadas
        for i,grupo1 in enumerate(verificar_salas_fixadas):
            # Salvo as salas repetidas, isto é, se houver duas aulas da mesma categoria de tamanho fixadas na mesma sala, a sala é salva
            # na variável salas_repetidas.
            # Ex: Suponha que existam duas aulas com cerca de 40 alunos que possuem conflito de horário,
            # e ambas foram fixadas na sala 3-009, então há um óbvio problema de que duas aulas estão fixadas no mesmo horário.
            salas_repetidas = [s for s in grupo1 if grupo1.count(s) > 1]
            # Se este caso for verdadeiro, isto é, duas aulas da mesma categoria estarem alocadas na mesma sala,
            # envio uma mensagem de erro, aponto quais aulas estão causando problema, e interrompo o código.
            if salas_repetidas:
                aula_repetida = [df.loc[int(aula % lenT), 'Disciplina (código)'] for aula in verificar_aulas[i]]
                aux = [a for a in aula_repetida if aula_repetida.count(a) > 1]
                if not aux:
                    print(f"Há aulas com conflito de horário que estão fixadas na mesma sala.")
                    print(f"Em particular, essas disciplinas que parecem estar dando problema:")
                    for aula in verificar_aulas[i]:
                        print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                    print("O grupo inteiro das aulas e disciplinas causando problema é esse:")
                    for grupinho in verificar_aulas:
                        for aula in grupinho:
                            print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")

                    custom_exit()

            # Caso não tenham salas fixadas para uma mesma categoria, passo a verificar se o mesmo não acontece com categorias diferentes.
            # Ex: Suponha que existam duas aulas com conflito de horário, que foram fixadas em uma mesma sala do bloco 5,
            # uma com 70 alunos e outra com 40 alunos. Há um claro problema na situação, quase que o mesmo anterior.
            # Portanto, para cada lista em verificar_salas_fixadas, isto é, para cada lista de salas fixadas
            for j,grupo2 in enumerate(verificar_salas_fixadas):
                # Verifico se, para categorias diferentes, há aulas fixadas na mesma sala.
                if i != j and set(grupo1) & set(grupo2):
                    # Se este caso for verdadeiro, isto é, duas aulas de categorias diferentes estarem alocadas na mesma sala,
                    # envio uma mensagem de erro, aponto quais aulas estão causando problema, e interrompo o código.
                    print(
                        f"Há aulas com conflito de horário que estão fixadas na mesma sala."
                    )
                    print(f"Em particular, essas disciplinas que parecem estar dando problema:")

                    for aula in verificar_aulas[i]:
                        print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                    for aula in verificar_aulas[j]:
                        print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")

                    print("O grupo inteiro das aulas e disciplinas causando problema é esse:")
                    for grupinho in verificar_aulas:
                        for aula in grupinho:
                            print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")

                    custom_exit()

        aux_verificar_aulas = verificar_aulas.copy()
        aux_verificar_salas = verificar_salas.copy()
        for i in range(len(verificar_aulas)):
            grupo = aux_verificar_aulas[i].copy()
            
            for aula in grupo[:]:
                if sala_fixa[aula] != '0':
                    sala_fixada = salas.index[salas['Sala'] == sala_fixa[aula]].tolist()[0]
                    for s in aux_verificar_salas[:]:
                        if sala_fixada in s:
                            s.remove(sala_fixada)
                    grupo.remove(aula)
            aux_verificar_aulas[i] = grupo
            
        
        aulas_alocadas = 0
        for g in range(len(verificar_aulas)):
            if len(aux_verificar_aulas[g]) > len(aux_verificar_salas[g]) - aulas_alocadas and len(aux_verificar_aulas[g]):
                print(
                    f"Há muitas aulas com conflito de horário no seguinte grupo, "
                    f"então uma troca de horários pode ser necessária, ou a diminuição do número de vagas da disciplina."
                    f"\nVerifique se alguma dessas disciplinas não foi proibida de ser alocada em uma sala específica, "
                    f"pois a proibição de uma pode afetar a alocação de outra."
                )
                print("Em particular, essas disciplinas que parecem estar dando problema:")
                for aula in verificar_aulas[g]:
                    print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                print("O grupo inteiro das aulas e disciplinas causando problema é esse:")
                for grupinho in verificar_aulas:
                    for aula in grupinho:
                        print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                
                custom_exit()
            aulas_alocadas += len(aux_verificar_aulas[g])

        # Se não há aulas com conflito de horário fixadas em uma mesma sala, o próximo passo é verificar a existência de aulas
        # suficientes para um determinado grupo de conflitos.
        # A variável aulas_alocadas é um contador para "simular" as aulas alocadas. O funcionamento dessa etapa é puramente matemático,
        # analisando a alocação das aulas maiores para as menores. Intuitivamente, colocamos as aulas maiores nas salas maiores que
        # conseguem comportar as aulas menores. Com isso, no momento de alocar as salas menores, precisamos excluir as salas maiores
        # que "já foram alocadas", e estas são contabilizadas com a variável aulas_alocadas.
        aulas_alocadas = 0
        

def verificar_horarios_de_conflito_lab(grupos_de_conflitos, salas_de_aulas):
    # Para cada grupo de conflito
    for grupo in grupos_de_conflitos:

        # Eu separo as aulas conflitantes em grupos baseados em quais salas elas cabem
        # E também separo as salas que suportam os grupos correspondentes
        grupo30 = [] # Aulas com até 30 alunos
        salas30 = []
        grupo60 = [] # Aulas com 31 a 60 alunos
        salas60 = []

        # Vou definir algumas listas com os índices de algumas salas que comportam essas categorias.
        # lista2 = salas.index[salas['Sala'].isin(['1-004'])].tolist()
        lista1 = salas.index[salas['Sala'].isin(['1-004', '6-303', '6-304', '6-305', '6-306', '6-307'])].tolist()
        lista0 = salas.index[salas['Sala'].isin(['6-303/6-304', '6-305/6-306'])].tolist()

        
        # A variável verificar_salas_fixadas é uma Lista das listas de salas fixadas pelas aulas.
        # Ex: Se duas aulas foram alocadas em salas grandes, então a primeira sublista contém quais dessas salas foram alocadas.
        # Isso ajudará nos momentos em que uma aula pequena for fixada em uma sala grande, pois na hora de verificar a
        # demanda de salas grandes, o código contabiliza essa fixação, ou seja, se eu tiver três aulas que precisam de salas grandes
        # e uma aula pequena fixada em uma sala grande, mas tenho apenas três salas grandes, o código perceberá esse conflito e mandará
        # um aviso ao usuário de que alguma aula foi alocada em uma sala requisitada por outras aulas.
        verificar_salas_fixadas = [[],[]]

        # Para cada aula no grupo em análise
        for aula in grupo:
            # Verifico em qual categoria a aula se enquadra, adicionando-a no grupo e salvando as salas que comportam a aula.
            if tam_t[int(aula % lenT)] <= 30:
                grupo30.append(aula)
                salas30.extend([s for s in salas_de_aulas[grupos_de_conflitos.index(grupo)] if eta_as[aula,s] == 1])
                salas30 = list(dict.fromkeys(salas30))
                last_entered = 1
            elif tam_t[int(aula % lenT)] <= 60 and tam_t[int(aula % lenT)] > 30:
                grupo60.append(aula)
                salas60.extend([s for s in salas_de_aulas[grupos_de_conflitos.index(grupo)] if eta_as[aula,s] == 1])
                salas60 = list(dict.fromkeys(salas60))
                last_entered = 0
            else:
                # Se eu não consegui colocar aquela aula em nenhuma das salas disponíveis, novamente notifico o erro.
                print(f"Não há sala capaz de comportar a disciplina {df.loc[aula % lenT, 'Disciplina (código)']}.")
                
                custom_exit()
            # Se a aula possui uma sala fixada, preciso editar a lista de verificação.
            if sala_fixa[aula] != '0':
                
                if salas[salas['Sala'] == sala_fixa[aula]].index[0] in lista0:
                    
                    # Também adiciono o índice da sala fixada na lista de mesmo índice do último grupo que a aula foi colocada, ou seja,
                    # se uma aula que pertence ao grupo77 foi fixada em uma sala do bloco 4, eu adiciono a sala na segunda posição da
                    # lista verificar_salas_fixadas, então de [[],[],[],[],[]] foi para [[],[5],[],[],[]].
                    # verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == str(df.loc[int(aula % lenT), 'Sala'])].index[0])
                    verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == sala_fixa[aula]].index[0])
                # As demais condições são equivalentes no raciocínio.
                elif salas[salas['Sala'] == sala_fixa[aula]].index[0] in lista1:
                    verificar_salas_fixadas[last_entered].append(salas[salas['Sala'] == sala_fixa[aula]].index[0])

        # Formados os subgrupos de aulas conflitantes separados, vou analisar a existência de aulas disponíveis para este grupo.
        # Ex: Se eu tenho 3 aulas de mesmo horário, mas apenas 2 salas disponíveis para alocá-las, há um problema com o grupo.
        verificar_aulas = [grupo60, grupo30]
        verificar_salas = [salas60, salas30]
        # print(verificar_aulas)
        # print(verificar_salas)

        # Para cada lista em verificar_salas_fixadas, isto é, para cada lista de salas fixadas
        for i,grupo1 in enumerate(verificar_salas_fixadas):
            # Salvo as salas repetidas, isto é, se houver duas aulas da mesma categoria de tamanho fixadas na mesma sala, a sala é salva
            # na variável salas_repetidas.
            # Ex: Suponha que existam duas aulas com cerca de 40 alunos que possuem conflito de horário,
            # e ambas foram fixadas na sala 3-009, então há um óbvio problema de que duas aulas estão fixadas no mesmo horário.
            salas_repetidas = [s for s in grupo1 if grupo1.count(s) > 1]
            # Se este caso for verdadeiro, isto é, duas aulas da mesma categoria estarem alocadas na mesma sala,
            # envio uma mensagem de erro, aponto quais aulas estão causando problema, e interrompo o código.
            if salas_repetidas:
                aula_repetida = [df.loc[int(aula % lenT), 'Disciplina (código)'] for aula in verificar_aulas[i]]
                aux = [a for a in aula_repetida if aula_repetida.count(a) > 1]
                if not aux:
                    print(f"Há aulas com conflito de horário que estão fixadas na mesma sala.")
                    print(f"Em particular, essas disciplinas que parecem estar dando problema:")
                    for aula in verificar_aulas[i]:
                        print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                    print("O grupo inteiro das aulas e disciplinas causando problema é esse:")
                    for grupinho in verificar_aulas:
                        for aula in grupinho:
                            print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                    
                    custom_exit()

            # Caso não tenham salas fixadas para uma mesma categoria, passo a verificar se o mesmo não acontece com categorias diferentes.
            # Ex: Suponha que existam duas aulas com conflito de horário, que foram fixadas em uma mesma sala do bloco 5,
            # uma com 70 alunos e outra com 40 alunos. Há um claro problema na situação, quase que o mesmo anterior.
            # Portanto, para cada lista em verificar_salas_fixadas, isto é, para cada lista de salas fixadas
            for j,grupo2 in enumerate(verificar_salas_fixadas):
                # Verifico se, para categorias diferentes, há aulas fixadas na mesma sala.
                if i != j and set(grupo1) & set(grupo2):
                    # Se este caso for verdadeiro, isto é, duas aulas de categorias diferentes estarem alocadas na mesma sala,
                    # envio uma mensagem de erro, aponto quais aulas estão causando problema, e interrompo o código.
                    print(
                        f"Há aulas com conflito de horário que estão fixadas na mesma sala."
                    )
                    print(f"Em particular, essas disciplinas que parecem estar dando problema:")

                    for aula in verificar_aulas[i]:
                        print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                    for aula in verificar_aulas[j]:
                        print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")

                    print("O grupo inteiro das aulas e disciplinas causando problema é esse:")
                    for grupinho in verificar_aulas:
                        for aula in grupinho:
                            print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                    
                    custom_exit()

        aux_verificar_aulas = verificar_aulas.copy()
        aux_verificar_salas = verificar_salas.copy()
        for i in range(len(verificar_aulas)):
            grupo = aux_verificar_aulas[i].copy()
            
            for aula in grupo[:]:
                if sala_fixa[aula] != '0':
                    sala_fixada = salas.index[salas['Sala'] == sala_fixa[aula]].tolist()[0]
                    for s in aux_verificar_salas[:]:
                        if sala_fixada in s:
                            s.remove(sala_fixada)
                    grupo.remove(aula)
            aux_verificar_aulas[i] = grupo
        aulas_alocadas = 0
        for g in range(len(verificar_aulas)):
            if len(aux_verificar_aulas[g]) > len(aux_verificar_salas[g]) - aulas_alocadas and len(aux_verificar_aulas[g]):
                print(
                    f"Há muitas aulas com conflito de horário no seguinte grupo, "
                    f"então uma troca de horários pode ser necessária, ou a diminuição do número de vagas da disciplina."
                    f"\nVerifique se alguma dessas disciplinas não foi proibida de ser alocada em uma sala específica, "
                    f"pois a proibição de uma pode afetar a alocação de outra."
                )
                print("Em particular, essas disciplinas que parecem estar dando problema:")
                for aula in verificar_aulas[g]:
                    print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                print("O grupo inteiro das aulas e disciplinas causando problema é esse:")
                for grupinho in verificar_aulas:
                    for aula in grupinho:
                        print(f"Aula {aula}, {df['Disciplina (código)'][int(aula % lenT)]}")
                
                custom_exit()
            aulas_alocadas += len(aux_verificar_aulas[g])
        # Se não há aulas com conflito de horário fixadas em uma mesma sala, o próximo passo é verificar a existência de aulas
        # suficientes para um determinado grupo de conflitos.
        # A variável aulas_alocadas é um contador para "simular" as aulas alocadas. O funcionamento dessa etapa é puramente matemático,
        # analisando a alocação das aulas maiores para as menores. Intuitivamente, colocamos as aulas maiores nas salas maiores que
        # conseguem comportar as aulas menores. Com isso, no momento de alocar as salas menores, precisamos excluir as salas maiores
        # que "já foram alocadas", e estas são contabilizadas com a variável aulas_alocadas.
        aulas_alocadas = 0
        


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

# Verificação de horários conflitantes.
# Se, por algum motivo, tiver mais grupos conflitantes que de salas para esses grupos, isso deve indicar que
# existem aulas que não estão em conflito com nenhuma outra, mas não conseguem ser alocadas em nenhuma sala.
# Esse erro já deveria ter acontecido na construção desses grupos, então nem deveria chegar aqui.
# Caso chegue neste ponto, envio uma mensagem de erro e interrompo o código.
if len(grupos_de_conflitos) > len(salas_de_aula):
    print("Há um problema entre os conjuntos de aulas conflitantes e as salas disponíveis para estes conjuntos.")

    custom_exit()
else:
    # Caso contrário, faço uma análise das aulas que estão em conflito de horário umas com as outras
    verificar_horarios_de_conflito(grupos_de_conflitos, salas_de_aula)

if len(laboratorios_conflito) > len(salas_de_laboratorio):
    print("Há um problema entre os conjuntos de aulas de laboratório conflitantes e os laboratórios disponíveis para estes conjuntos.")
    custom_exit()
else:
    # Caso contrário, faço uma análise das aulas que estão em conflito de horário umas com as outras
    verificar_horarios_de_conflito_lab(laboratorios_conflito, salas_de_laboratorio)


# Se nenhuma das situações de erro foi encontrada, aviso que não há nenhum erro aparente com os dados fornecidos.
print("Aparentemente, não há nenhum conflito de horário, e há salas disponíveis em todos os horários necessários.")
