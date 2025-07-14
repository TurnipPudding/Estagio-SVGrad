import pandas as pd
import os
import sys
import traceback


# salas = pd.read_excel('C:/Users/gabri/Estágio/Dados/Dados_das_salas.xlsx', sheet_name="Salas")
# file_path = 'C:/Users/gabri/Estágio/Dados/Dados_das_salas_atualizado.xlsx'
# file_path = 'C:/Users/gabri/Estágio/Dados/Dados das salas 2025.xlsx'
df_completo = pd.ExcelFile(sys.argv[1])

sheet_names = df_completo.sheet_names
# salas = pd.read_excel(sys.argv[1], sheet_name="Salas")
salas = pd.read_excel(sys.argv[1], sheet_name=sheet_names[0])

# sheets = ["SME", "SMA", "SCC", "SSC"]

# df = pd.read_excel('C:/Users/gabri/Estágio/Dados/Dados_das_salas.xlsx', sheet_name=sheets)
# df = pd.read_excel(sys.argv[1], sheet_name=sheets)
df = pd.read_excel(sys.argv[1], sheet_name=sheet_names[1:])


# print(df)
# vagas = pd.read_excel('C:/Users/gabri/Estágio/Dados/planilhas jupiter 2025.xlsx', sheet_name=sheets)
# vagas = pd.read_excel('C:/Users/gabri/Estágio/Dados/planilhas_jupiter.xlsx', sheet_name=sheets)
jpter = pd.ExcelFile(sys.argv[2])
sheet_names_jpter = jpter.sheet_names
vagas = pd.read_excel(sys.argv[2], sheet_name=sheet_names_jpter[0:])

ingressantes = pd.read_excel(sys.argv[3])

espelho = pd.read_excel(sys.argv[4])

# if 'Disciplina (código)' not in df.columns:
#     msg = f"A coluna 'Disciplina (código)' não foi encontrada no arquivo {os.path.basename(sys.argv[1])}."
#     "Verifique o arquivo e tente novamente."
#     print(msg, file=sys.stderr)
#     sys.exit(4)
# elif 'Turma' not in df.columns:
#     msg = f"A coluna 'Turma' não foi encontrada no arquivo {os.path.basename(sys.argv[1])}."
#     "Verifique o arquivo e tente novamente."
#     print(msg, file=sys.stderr)
#     sys.exit(4)
# elif 'Disciplina' not in vagas.columns:
#     msg = f"A coluna 'Disciplina' não foi encontrada no arquivo {os.path.basename(sys.argv[2])}."
#     "Verifique o arquivo e tente novamente."
#     print(msg, file=sys.stderr)
#     sys.exit(4)

try:
    for d in range(len(ingressantes['Disciplina (código)'])):
        if ' ' in str(ingressantes.loc[d, 'Disciplina (código)']):
            ingressantes.loc[d, 'Disciplina (código)'] = str(ingressantes.loc[d, 'Disciplina (código)']).replace(' ', '')
        if '-' not in str(ingressantes.loc[d, 'Disciplina (código)']):
            ingressantes.loc[d, 'Disciplina (código)'] = \
            f"{ingressantes.loc[d, 'Disciplina (código)']}-{int(ingressantes.loc[d, 'Turma'])}"
except KeyError as e:
    coluna = str(e).strip('\'')
    msg = f"A coluna '{coluna}' não foi encontrada no arquivo {os.path.basename(sys.argv[3])}. Verifique o arquivo."
    print(msg, file=sys.stderr)
    sys.exit(4)

try:
    for d in range(len(espelho['Disciplina (código)'])):
        if ' ' in str(espelho.loc[d, 'Disciplina (código)']):
            espelho.loc[d, 'Disciplina (código)'] = str(espelho.loc[d, 'Disciplina (código)']).replace(' ', '')
        if '-' not in str(espelho.loc[d, 'Disciplina (código)']):
            espelho.loc[d, 'Disciplina (código)'] = \
            f"{espelho.loc[d, 'Disciplina (código)']}-{int(espelho.loc[d, 'Turma'])}"
except KeyError as e:
    coluna = str(e).strip('\'')
    msg = f"A coluna '{coluna} não foi encontrada no arquivo {os.path.basename(sys.argv[4])}. Verifique o arquivo."
    print(msg, file=sys.stderr)
    sys.exit(4)

for sheet in sheet_names_jpter[0:]:
    vagas[sheet] = vagas[sheet].fillna(0)

# df = pd.read_excel('C:/Users/gabri/Estágio/Dados/Dados_das_salas_copia.xlsx', sheet_name=sheets)

# Para as planilhas de SME, SMA, SCC e SSC (4 primeiras planilhas depois da de salas)
for sheet in sheet_names[1:5]:
    df[sheet].columns = [col.replace("\n", " ") for col in df[sheet].columns]
    
    try:
        for d in range(len(vagas[sheet]['Disciplina'])):
            vagas[sheet].loc[d, 'Disciplina'] = \
            f"{vagas[sheet].loc[d, 'Disciplina']}-{int(vagas[sheet].loc[d, 'Turma'] % 100)}"
    except KeyError as e:
        coluna = str(e).strip('\'')
        msg = f"A coluna '{coluna} não foi encontrada no arquivo {os.path.basename(sys.argv[2])}. Verifique o arquivo."
        print(msg, file=sys.stderr)
        sys.exit(4)
        
    try:
        for d in range(len(df[sheet]['Disciplina (código)'])):
            if ' ' in str(df[sheet].loc[d, 'Disciplina (código)']):
                df[sheet].loc[d, 'Disciplina (código)'] = str(df[sheet].loc[d, 'Disciplina (código)']).replace(' ', '')
            if '-' not in str(df[sheet].loc[d, 'Disciplina (código)']):
                df[sheet].loc[d, 'Disciplina (código)'] = \
                f"{df[sheet].loc[d, 'Disciplina (código)']}-{int(df[sheet].loc[d, 'Turma'])}"
    except KeyError as e:
        coluna = str(e).strip('\'')
        msg = f"A coluna '{coluna} não foi encontrada no arquivo {os.path.basename(sys.argv[1])}. Verifique o arquivo."
        print(msg, file=sys.stderr)
        sys.exit(4)

    # print(vagas['SME'])
    # print(df['SME'])
    colunas = vagas[sheet].columns
    # lista1 = ["Vagas obrigatórias", "Vagas eletivas", "Vagas optativas livres", "Vagas especiais", "Vagas extras"]
    lista1 = ["Vagas obrigatórias", "Vagas eletivas", "Vagas optativas livres", "Vagas especiais"]
    lista2 = []
    for l in lista1:
        try:
            lista2.append(colunas.get_loc(l) + 1)
        except KeyError as e:
            coluna = str(e).strip('\'')
            msg = f"A coluna '{coluna} não foi encontrada no arquivo {os.path.basename(sys.argv[2])}. "
            f"Verifique a planilha {sheet}."
            print(msg, file=sys.stderr)
            sys.exit(4)
    for d in range(len(df[sheet]['Disciplina (código)'])):
        # if df[sheet].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
        #     index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
        #     df[sheet].loc[d, 'Vagas por disciplina'] = vagas[sheet].loc[index, 'Inscritos obrigatórios'] + \
        #     vagas[sheet].loc[index, 'Inscritos eletivos'] + vagas[sheet].loc[index, 'Inscritos livres'] + \
        #     vagas[sheet].loc[index, 'Inscritos extras'] + vagas[sheet].loc[index, 'Inscritos especiais']
        # else:
        #     df[sheet].loc[d, 'Vagas por disciplina'] = 0
        # if df[sheet].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
        #     index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
        #     # vagas[sheet].loc[index, coluna[lista2[0]]]
        #     df[sheet].loc[d, 'Vagas por disciplina'] = vagas[sheet].loc[index, colunas[lista2[0]]] + \
        #     vagas[sheet].loc[index, colunas[lista2[1]]] + vagas[sheet].loc[index, colunas[lista2[2]]] + \
        #     vagas[sheet].loc[index, colunas[lista2[3]]] + vagas[sheet].loc[index, colunas[lista2[4]]]
        if df[sheet].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
            index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
            # vagas[sheet].loc[index, coluna[lista2[0]]]
            df[sheet].loc[d, 'Vagas por disciplina'] = 0
            for i in range(len(lista2)):
                df[sheet].loc[d, 'Vagas por disciplina'] += vagas[sheet].loc[index, colunas[lista2[i]]]
            # df[sheet].loc[d, 'Vagas por disciplina'] = vagas[sheet].loc[index, colunas[lista2[0]]] + \
            # vagas[sheet].loc[index, colunas[lista2[1]]] + vagas[sheet].loc[index, colunas[lista2[2]]] + \
            # vagas[sheet].loc[index, colunas[lista2[3]]]
        else:
            df[sheet].loc[d, 'Vagas por disciplina'] = 0

        if pd.isna(df[sheet].loc[d, 'Observações']) or df[sheet].loc[d, 'Observações'] != 0:
            if "Ingressantes" in str(df[sheet].loc[d, 'Observações']):
                index = (ingressantes['Disciplina (código)'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
                df[sheet].loc[d, 'Vagas por disciplina'] += ingressantes.loc[index, 'Ingressantes']

        if pd.isna(df[sheet].loc[d, 'Observações']) or df[sheet].loc[d, 'Observações'] != 0:
            if "Espelho" in str(df[sheet].loc[d, 'Observações']):
                index = (espelho['Disciplina (código)'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
                df[sheet].loc[d, 'Vagas por disciplina'] += espelho.loc[index, 'Inscritos']
            
            
            
    
        
# df = pd.read_excel(sys.argv[1], sheet_name="Outros")
# vagas = pd.read_excel(sys.argv[2], sheet_names_jpter[4:])
df[sheet_names[5]].columns = [col.replace("\n", " ") for col in df[sheet_names[5]].columns]

for d in range(len(df[sheet_names[5]]['Disciplina (código)'])):
    try:
        if ' ' in str(df[sheet_names[5]].loc[d, 'Disciplina (código)']):
            df[sheet_names[5]].loc[d, 'Disciplina (código)'] = str(df[sheet_names[5]].loc[d, 'Disciplina (código)']).replace(' ', '')
        if '-' not in str(df[sheet_names[5]].loc[d, 'Disciplina (código)']):
            df[sheet_names[5]].loc[d, 'Disciplina (código)'] = \
            f"{df[sheet_names[5]].loc[d, 'Disciplina (código)']}-{int(df[sheet_names[5]].loc[d, 'Turma'])}"
    except KeyError as e:
        coluna = str(e).strip('\'')
        msg = f"A coluna '{coluna} não foi encontrada no arquivo {os.path.basename(sys.argv[1])}. Verifique o arquivo."
        print(msg, file=sys.stderr)
        sys.exit(4)
        
for sheet in sheet_names_jpter[4:]:
    try:
        for d in range(len(vagas[sheet]['Disciplina'])):
            vagas[sheet].loc[d, 'Disciplina'] = \
            f"{vagas[sheet].loc[d, 'Disciplina']}-{int(vagas[sheet].loc[d, 'Turma'] % 100)}"
    except KeyError as e:
        coluna = str(e).strip('\'')
        msg = f"A coluna '{coluna} não foi encontrada no arquivo {os.path.basename(sys.argv[2])}. Verifique o arquivo."
        print(msg, file=sys.stderr)
        sys.exit(4)
        
    colunas = vagas[sheet].columns
    # lista1 = ["Vagas obrigatórias", "Vagas eletivas", "Vagas optativas livres", "Vagas especiais", "Vagas extras"]
    lista1 = ["Vagas obrigatórias", "Vagas eletivas", "Vagas optativas livres", "Vagas especiais"]
    lista2 = []
    for l in lista1:
        try:
            lista2.append(colunas.get_loc(l) + 1)
        except KeyError as e:
            coluna = str(e).strip('\'')
            msg = f"A coluna '{coluna}' não foi encontrada no arquivo {os.path.basename(sys.argv[2])}. "
            f"Verifique a planilha {sheet}."
            print(msg, file=sys.stderr)
            sys.exit(4)
    # print(vagas[sheet]['Disciplina'])
    for d in range(len(df[sheet_names[5]]['Disciplina (código)'])):
        # print(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
        
        # if df[sheet_names[5]].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
        #     index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
        #     df[sheet_names[5]].loc[d, 'Vagas por disciplina'] = int(vagas[sheet].loc[index, colunas[lista2[0]]]) + \
        #     int(vagas[sheet].loc[index, colunas[lista2[1]]]) + int(vagas[sheet].loc[index, colunas[lista2[2]]]) + \
        #     int(vagas[sheet].loc[index, colunas[lista2[3]]]) + int(vagas[sheet].loc[index, colunas[lista2[4]]])
        if df[sheet_names[5]].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
            index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])

            df[sheet_names[5]].loc[d, 'Vagas por disciplina'] = 0
            for i in range(len(lista2)):
                df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += vagas[sheet].loc[index, colunas[lista2[i]]]
            # df[sheet_names[5]].loc[d, 'Vagas por disciplina'] = int(vagas[sheet].loc[index, colunas[lista2[0]]]) + \
            # int(vagas[sheet].loc[index, colunas[lista2[1]]]) + int(vagas[sheet].loc[index, colunas[lista2[2]]]) + \
            # int(vagas[sheet].loc[index, colunas[lista2[3]]])
        elif pd.isna(df[sheet_names[5]].loc[d, 'Vagas por disciplina']):
            df[sheet_names[5]].loc[d, 'Vagas por disciplina'] = 0
            
for d in range(len(df[sheet_names[5]]['Disciplina (código)'])):            
    if pd.isna(df[sheet_names[5]].loc[d, 'Observações']) or df[sheet_names[5]].loc[d, 'Observações'] != 0:
        if "Ingressantes" in str(df[sheet_names[5]].loc[d, 'Observações']):
            index = (ingressantes['Disciplina (código)'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
            df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += ingressantes.loc[index, 'Ingressantes']
    if pd.isna(df[sheet_names[5]].loc[d, 'Observações']) or df[sheet_names[5]].loc[d, 'Observações'] != 0:
            if "Espelho" in str(df[sheet_names[5]].loc[d, 'Observações']):
                index = (espelho['Disciplina (código)'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
                df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += espelho.loc[index, 'Inscritos']
            

# curriculos = ['BMACC', 'BMA', 'LMA', 'MAT-NG', 'BECD', 'BCC', 'BSI', 'BCDados']
# ingressantes = pd.read_excel('C:/Users/gabri/Estágio/Códigos/Endgame/Testes/Dados dos ingressantes.xlsx')
# for sheet in sheet_names[1:]:
#     for d in range(len(df[sheet])):
#         # print(f"Valor de tam_t: {tam_t[d]}")
#         # print(f"Valor do dataframe: {df.loc[d, 'Vagas por disciplina']}")
#         if not pd.isna(df[sheet].loc[d, 'Observações']):
#             if 'Ingressantes' in df[sheet].loc[d, 'Observações'] or 'ingressantes' in df[sheet].loc[d, 'Observações']:
#                 print(
#                     f"Número de inscritos na disciplina {df[sheet].loc[d, 'Disciplina (código)']} (disciplina de ingressantes): {df[sheet].loc[d, 'Vagas por disciplina']}"
#                     f"\nAdicionando número de ingressantes fornecido pelo usuário."
#                 )
#                 # print(df.loc[d, 'Vagas por disciplina'])
#                 # print(tam_t[d])
#                 if ',' in df[sheet].loc[d, 'Observações']:
#                     lista = df[sheet].loc[d, 'Curso(s)'].split(", ")
#                     for c in lista:
#                         df[sheet].loc[d, 'Vagas por disciplina'] += ingressantes.loc[ingressantes.index[ingressantes['Curso'] == c], 'Número de ingressantes']
#                 else:
#                     c = df[sheet].loc[d, 'Curso(s)']
#                     df[sheet].loc[d, 'Vagas por disciplina'] += ingressantes.loc[ingressantes.index[ingressantes['Curso'] == c], 'Número de ingressantes']
#                 # tam_t[d] += qtd_pos
#                 # df.loc[d, 'Vagas por disciplina'] += qtd_pos
#                 print(
#                     f"Número de inscritos na disciplina {df[sheet].loc[d, 'Disciplina (código)']} (disciplina de ingressantes): {df[sheet].loc[d, 'Vagas por disciplina']}"
#                 )
# file_path = 'C:/Users/gabri/Estágio/Dados/Dados das salas 2024 copia.xlsx'
full_name = sys.argv[5]
file_path = os.path.join(os.getcwd(), "Saídas da Interface", "Planilhas de Dados", full_name)

# if os.path.exists(file_path):
#     os.remove(file_path)
#     print(f"Arquivo existente '{file_path}' removido.")

    
dfs = {
    "Salas": salas,
    "SME": df['SME'],
    "SMA": df['SMA'],
    "SCC": df['SCC'],
    "SSC": df['SSC'],
    "Outros": df['Outros']
}

try:
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

except PermissionError as e:
    if e.errno == 13:  # Erro de permissão (arquivo aberto ou bloqueado)
        traceback.print_exc(file=sys.stderr)
        sys.exit(2)
        
    else:
        traceback.print_exc(file=sys.stderr)
        sys.exit(3)
        
except Exception as e:
    # Para qualquer outro erro
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)