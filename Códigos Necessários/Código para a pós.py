import pandas as pd
from datetime import datetime

df_livres = pd.read_excel('C:/Users/gabri/Estágio/Códigos/Demonstração/Saídas da Interface/Planilhas de Dados/plan1.xlsx')


# # Função auxiliar para converter horário "HH:MM" em datetime.time
# def str_to_time(horastr):
#     return datetime.strptime(horastr, "%H:%M").time()

# # Pré-processamento: transformar os intervalos livres em objetos time
# intervalos_livres = {}  # chave: (sala, dia), valor: lista de tuplas (inicio, fim)
# for idx, row in df_livres.iterrows():
#     sala = row['Sala']
#     dia = row['Dia da semana']
#     inicio_str, fim_str = str(row['Horário vago']).split(' - ')
#     inicio = str_to_time(inicio_str.strip())
#     fim = str_to_time(fim_str.strip())
#     chave = (sala, dia)
#     intervalos_livres.setdefault(chave, []).append((inicio, fim))




file_name = "C:/Users/gabri/Estágio/Códigos/Demonstração/Files/Dados da Pós/Elenco CCMC 202501.xlsx"
# Leio e salvo o arquivo em uma variável.
df_pos = pd.read_excel(file_name)

# Defino uma variável com o nome de um cabeçalho para ser encontrado, caso o cabeçalho não seja a primeira linha da planilha.
header_name = 'Disciplina (código)'

header_found = False
# Para cada linha e célula da primeira coluna do dataframe:
for i, valor in enumerate(df_pos.loc[:,df_pos.columns[0]]):
    # Se o valor da célula for o nome do cabeçalho que estou procurando:
    if valor == header_name:
        # Salvo o número da linha do cabeçalho.
        header_row = i+1
        df_pos = pd.read_excel(file_name, header=header_row)
        header_found = True
        break

print(df_pos.columns)


# Função para converter horário no formato 'HH:MM' para valor decimal em horas
def horario_para_decimal(horario):
    # Se, por algum acaso, houver um horário definido como 20h40, ele é convertido para 20:40
    if 'h' in horario:
        horario = horario.replace('h',':')
    # Separo e identifico os componentes daquele horário, ou seja, salvo o valor de horas e o de minutos
    horas, minutos = map(int, horario.split(':'))
    # Retorno um valor numérico daquele horário seguindo a ideia de "porcentagem" de hora.
    # Ex: 40 minutos são dois terços de uma hora (40/60 = 2/3), logo, 20:40 pode ser traduzido como 20 + 40/60 = 20.67
    return horas + minutos / 60


# Função para processar a célula no formato 'Dia - HH:MM/HH:MM'
def processar_horario(celula):
#     print(celula)
    # Verifico se a célula que está sendo analisada possui um horário definido (se há algo escrito nela e separado por um traço "-")
    if isinstance(celula, str) and "-" in celula:
        # Deleto qualquer espaço " " da célula para garantir uma leitura organizada,
        # já que algumas vezes pode haver mais de um espaço, ou a ausência dos mesmos
        celula = str(celula).replace(' ', '')
        # print(celula)
        # Separo e salvo o dia da semana o qual a aula é dada, e o horário do dia que ela é ministrada
        dia, horarios = celula.split('-')
        # Separo o horário de início e de término daquela aula
        inicio, fim = horarios.split('/')
        # Salvo o horário do ínicio e do fim daquela aula
        start_a = horario_para_decimal(inicio)
        end_a = horario_para_decimal(fim)
        # Retorno o dia e os horários daquela aula
        return dia, start_a, end_a
    else:
        # Se a célula é vazia ou não possui o traço "-", é uma célula irregular com horário não definido, logo, retorno valor 0
        return 0, 0, 0

# Lista das colunas de horários
colunas_horarios = ['Horário 1', 'Horário 2', 'Horário 3', 'Horário 4']
result = []
# Processar cada coluna de horários
# Para cada coluna de horários
for coluna in colunas_horarios:
    # Aplico o processamento e tradução do horário
    resultados = df_pos[coluna].apply(processar_horario).to_list()
    # Adiciono numa lista os horários traduzidos de cada coluna
    result.extend(resultados)

# Crio um dataframe com os dados padronizados de todas as aulas
A = pd.DataFrame(result, columns=['Dia', 'start_a', 'end_a'])
# Salvo as colunas do dataframe em listas separadas
dia_a = A['Dia'].to_list()
start_a = A['start_a'].to_list()
end_a = A['end_a'].to_list()

print(dia_a)
print(start_a)
print(end_a)

file_path = "C:/Users/gabri/Estágio/Códigos/Demonstração/Files/Interface/Planilha das salas.xlsx"
salas = pd.read_excel(file_path, sheet_name="Salas")
tam_t = df_pos['Vagas por disciplina'].tolist()
cap_s = salas['Lugares'].tolist()

lenT = len(range(len(df_pos['Disciplina (código)'])))
lenS = len(range(len(salas['Sala'])))
lenA = len(A)

eta_as = {(a, s): 1 if tam_t[int((a % lenT ))] <= cap_s[s] \
          else 0 for a in range(lenA) for s in range(lenS)}

print(eta_as)

# Preciso usar o dataframe df_livres para verificar os horários livres de cada sala
# No df_livres, cada linha tem uma sala, e essa sala tem um index no eta_as.
# Além disso, na mesma linha, há um dia da semana e um horário vago.
# Então, para cada linha do df_livres, eu preciso verificar se o horário daquela sala está livre
for idx, row in df_livres.iterrows():
    sala = row['Sala']
    dia = row['Dia da semana']
    inicio_str, fim_str = str(row['Horário vago']).split(' - ')
    inicio = horario_para_decimal(inicio_str.strip())
    fim = horario_para_decimal(fim_str.strip())
    
    # Verifico se a sala está livre para cada aula
    for a in range(lenA):
        # Se o dia da aula for o mesmo que o dia do horário livre
        # e o horário da aula não estiver completamente dentro do horário livre,
        if dia_a[a] == dia and not (start_a[a] >= inicio and end_a[a] <= fim):
            # Verifico se a aula caberia na sala
            s = salas[salas['Sala'] == sala].index[0]
            if eta_as[(a, s)] == 1:
                # Se a aula caberia na sala, marco como 0 (não cabe),
                # pois o horário não está livre para aquela aula
                eta_as[(a, s)] = 0

print(eta_as)