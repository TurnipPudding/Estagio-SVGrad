
# Script para processamento e padronização de dados de disciplinas, vagas e ingressantes.
# Lê arquivos Excel fornecidos via linha de comando, trata códigos de disciplina e prepara DataFrames para integração.
# Realiza validação de colunas críticas e padroniza formatos para garantir consistência entre planilhas.

import pandas as pd
import os
import sys
import traceback



# Lê o arquivo principal com todas as planilhas (salas e disciplinas)
df_completo = pd.ExcelFile(sys.argv[1])
# Extrai os nomes das planilhas presentes no arquivo
sheet_names = df_completo.sheet_names
# Lê a planilha de salas (primeira aba)
salas = pd.read_excel(sys.argv[1], sheet_name=sheet_names[0])
# Lê as planilhas de disciplinas (demais abas)
df = pd.read_excel(sys.argv[1], sheet_name=sheet_names[1:])

# Lê o arquivo de vagas do JúpiterWeb
jpter = pd.ExcelFile(sys.argv[2])
sheet_names_jpter = jpter.sheet_names
# Lê todas as planilhas de vagas
vagas = pd.read_excel(sys.argv[2], sheet_name=sheet_names_jpter[0:])

# Lê planilha de ingressantes (novos alunos)
ingressantes = pd.read_excel(sys.argv[3])
# Lê planilha de inscritos de disciplinas espelho
espelho = pd.read_excel(sys.argv[4])


# Padroniza os códigos das disciplinas dos ingressantes, removendo espaços e adicionando o sufixo da turma se necessário
try:
    for d in range(len(ingressantes['Disciplina (código)'])):
        # Remove espaços do código da disciplina
        if ' ' in str(ingressantes.loc[d, 'Disciplina (código)']):
            ingressantes.loc[d, 'Disciplina (código)'] = str(ingressantes.loc[d, 'Disciplina (código)']).replace(' ', '')
        # Adiciona sufixo da turma se não houver '-'
        if '-' not in str(ingressantes.loc[d, 'Disciplina (código)']):
            ingressantes.loc[d, 'Disciplina (código)'] = \
            f"{ingressantes.loc[d, 'Disciplina (código)']}-{int(ingressantes.loc[d, 'Turma'])}"
except KeyError as e:
    # Se faltar coluna, exibe mensagem de erro detalhada e encerra o script
    coluna = str(e).strip('\'')
    msg = f"A coluna '{coluna}' não foi encontrada no arquivo {os.path.basename(sys.argv[3])}. Verifique o arquivo."
    print(msg, file=sys.stderr)
    sys.exit(4)

# Padroniza os códigos das disciplinas do espelho, removendo espaços e adicionando o sufixo da turma se necessário
try:
    for d in range(len(espelho['Disciplina (código)'])):
        if ' ' in str(espelho.loc[d, 'Disciplina (código)']):
            espelho.loc[d, 'Disciplina (código)'] = str(espelho.loc[d, 'Disciplina (código)']).replace(' ', '')
        if '-' not in str(espelho.loc[d, 'Disciplina (código)']):
            espelho.loc[d, 'Disciplina (código)'] = \
            f"{espelho.loc[d, 'Disciplina (código)']}-{int(espelho.loc[d, 'Turma'])}"
except KeyError as e:
    # Se faltar coluna, exibe mensagem de erro detalhada e encerra o script
    coluna = str(e).strip('\'')
    msg = f"A coluna '{coluna}' não foi encontrada no arquivo {os.path.basename(sys.argv[4])}. Verifique o arquivo."
    print(msg, file=sys.stderr)
    sys.exit(4)

# Preenche valores NaN com zero nas planilhas de vagas do JúpiterWeb para evitar problemas em cálculos posteriores
for sheet in sheet_names_jpter[0:]:
    vagas[sheet] = vagas[sheet].fillna(0)


# Processa as planilhas de SME, SMA, SCC e SSC (as 4 primeiras após a de salas)
for sheet in sheet_names[1:5]:
    # Remove quebras de linha dos nomes das colunas para evitar problemas de acesso
    df[sheet].columns = [col.replace("\n", " ") for col in df[sheet].columns]

    # Padroniza o código das disciplinas nas planilhas de vagas, adicionando sufixo da turma
    try:
        for d in range(len(vagas[sheet]['Disciplina'])):
            vagas[sheet].loc[d, 'Disciplina'] = \
            f"{vagas[sheet].loc[d, 'Disciplina']}-{int(vagas[sheet].loc[d, 'Turma'] % 100)}"
    except KeyError as e:
        coluna = str(e).strip('\'')
        msg = f"A coluna '{coluna}' não foi encontrada no arquivo {os.path.basename(sys.argv[2])}. Verifique o arquivo."
        print(msg, file=sys.stderr)
        sys.exit(4)

    # Padroniza o código das disciplinas nas planilhas de disciplinas, removendo espaços e adicionando sufixo da turma
    try:
        for d in range(len(df[sheet]['Disciplina (código)'])):
            if ' ' in str(df[sheet].loc[d, 'Disciplina (código)']):
                df[sheet].loc[d, 'Disciplina (código)'] = str(df[sheet].loc[d, 'Disciplina (código)']).replace(' ', '')
            if '-' not in str(df[sheet].loc[d, 'Disciplina (código)']):
                df[sheet].loc[d, 'Disciplina (código)'] = \
                f"{df[sheet].loc[d, 'Disciplina (código)']}-{int(df[sheet].loc[d, 'Turma'])}"
    except KeyError as e:
        coluna = str(e).strip('\'')
        msg = f"A coluna '{coluna}' não foi encontrada no arquivo {os.path.basename(sys.argv[1])}. Verifique o arquivo."
        print(msg, file=sys.stderr)
        sys.exit(4)

    # Identifica as colunas de vagas relevantes
    colunas = vagas[sheet].columns
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

    # Calcula o total de vagas por disciplina, somando todas as categorias relevantes
    for d in range(len(df[sheet]['Disciplina (código)'])):
        # Se a disciplina está presente na planilha de vagas, soma as vagas das colunas relevantes
        if df[sheet].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
            index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
            df[sheet].loc[d, 'Vagas por disciplina'] = 0
            for i in range(len(lista2)):
                df[sheet].loc[d, 'Vagas por disciplina'] += vagas[sheet].loc[index, colunas[lista2[i]]]
        else:
            # Se não está presente, define como zero
            df[sheet].loc[d, 'Vagas por disciplina'] = 0

        # Se há observação de ingressantes, soma o número de ingressantes à vaga
        if pd.isna(df[sheet].loc[d, 'Observações']) or df[sheet].loc[d, 'Observações'] != 0:
            if "Ingressantes" in str(df[sheet].loc[d, 'Observações']):
                if df[sheet].loc[d, 'Disciplina (código)'] in ingressantes['Disciplina (código)'].tolist():
                    index = (ingressantes['Disciplina (código)'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
                    df[sheet].loc[d, 'Vagas por disciplina'] += ingressantes.loc[index, 'Ingressantes']

        # Se há observação de espelho, soma o número de inscritos do espelho à vaga
        if pd.isna(df[sheet].loc[d, 'Observações']) or df[sheet].loc[d, 'Observações'] != 0:
            if "Espelho" in str(df[sheet].loc[d, 'Observações']):
                if df[sheet].loc[d, 'Disciplina (código)'] in espelho['Disciplina (código)'].tolist():
                    index = (espelho['Disciplina (código)'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
                    df[sheet].loc[d, 'Vagas por disciplina'] += espelho.loc[index, 'Inscritos']
            


# Processa a planilha 'Outros' (sheet_names[5]), removendo quebras de linha dos nomes das colunas
df[sheet_names[5]].columns = [col.replace("\n", " ") for col in df[sheet_names[5]].columns]

# Padroniza o código das disciplinas na planilha 'Outros', removendo espaços e adicionando sufixo da turma
for d in range(len(df[sheet_names[5]]['Disciplina (código)'])):
    try:
        if ' ' in str(df[sheet_names[5]].loc[d, 'Disciplina (código)']):
            df[sheet_names[5]].loc[d, 'Disciplina (código)'] = str(df[sheet_names[5]].loc[d, 'Disciplina (código)']).replace(' ', '')
        if '-' not in str(df[sheet_names[5]].loc[d, 'Disciplina (código)']):
            df[sheet_names[5]].loc[d, 'Disciplina (código)'] = \
            f"{df[sheet_names[5]].loc[d, 'Disciplina (código)']}-{int(df[sheet_names[5]].loc[d, 'Turma'])}"
    except KeyError as e:
        coluna = str(e).strip('\'')
        msg = f"A coluna '{coluna}' não foi encontrada no arquivo {os.path.basename(sys.argv[1])}. Verifique o arquivo."
        print(msg, file=sys.stderr)
        sys.exit(4)

# Processa as planilhas de vagas extras (sheet_names_jpter[4:])
for sheet in sheet_names_jpter[4:]:
    # Padroniza o código das disciplinas nas planilhas de vagas extras
    try:
        for d in range(len(vagas[sheet]['Disciplina'])):
            vagas[sheet].loc[d, 'Disciplina'] = \
            f"{vagas[sheet].loc[d, 'Disciplina']}-{int(vagas[sheet].loc[d, 'Turma'] % 100)}"
    except KeyError as e:
        coluna = str(e).strip('\'')
        msg = f"A coluna '{coluna}' não foi encontrada no arquivo {os.path.basename(sys.argv[2])}. Verifique o arquivo."
        print(msg, file=sys.stderr)
        sys.exit(4)

    # Identifica as colunas de vagas relevantes
    colunas = vagas[sheet].columns
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

    # Calcula o total de vagas por disciplina na planilha 'Outros', somando todas as categorias relevantes
    for d in range(len(df[sheet_names[5]]['Disciplina (código)'])):
        # Se a disciplina está presente na planilha de vagas, soma as vagas das colunas relevantes
        if df[sheet_names[5]].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
            index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
            df[sheet_names[5]].loc[d, 'Vagas por disciplina'] = 0
            for i in range(len(lista2)):
                df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += vagas[sheet].loc[index, colunas[lista2[i]]]
        elif pd.isna(df[sheet_names[5]].loc[d, 'Vagas por disciplina']):
            # Se não está presente ou está NaN, define como zero
            df[sheet_names[5]].loc[d, 'Vagas por disciplina'] = 0
            

# Ajusta o total de vagas por disciplina na planilha 'Outros' considerando ingressantes e espelho
for d in range(len(df[sheet_names[5]]['Disciplina (código)'])):            
    # Se há observação de ingressantes, soma o número de ingressantes à vaga
    if pd.isna(df[sheet_names[5]].loc[d, 'Observações']) or df[sheet_names[5]].loc[d, 'Observações'] != 0:
        if "Ingressantes" in str(df[sheet_names[5]].loc[d, 'Observações']):
            if df[sheet_names[5]].loc[d, 'Disciplina (código)'] in ingressantes['Disciplina (código)'].tolist():
                index = (ingressantes['Disciplina (código)'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
                df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += ingressantes.loc[index, 'Ingressantes']
    # Se há observação de espelho, soma o número de inscritos do espelho à vaga
    if pd.isna(df[sheet_names[5]].loc[d, 'Observações']) or df[sheet_names[5]].loc[d, 'Observações'] != 0:
        if "Espelho" in str(df[sheet_names[5]].loc[d, 'Observações']):
            if df[sheet_names[5]].loc[d, 'Disciplina (código)'] in espelho['Disciplina (código)'].tolist():
                index = (espelho['Disciplina (código)'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
                df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += espelho.loc[index, 'Inscritos']
            



# Define o nome e caminho do arquivo de saída para as planilhas geradas
full_name = sys.argv[5]
file_path = os.path.join(os.getcwd(), "Saídas da Interface", "Planilhas de Dados", full_name)

# Monta o dicionário de DataFrames para exportação, cada chave corresponde a uma aba do Excel
dfs = {
    "Salas": salas,
    "SME": df['SME'],
    "SMA": df['SMA'],
    "SCC": df['SCC'],
    "SSC": df['SSC'],
    "Outros": df['Outros']
}

# Exporta os DataFrames para o arquivo Excel, tratando erros de permissão e outros imprevistos
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
    # Para qualquer outro erro, retorna o erro para o console
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)