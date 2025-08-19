import pandas as pd
import os
import sys
import traceback


df_completo = pd.ExcelFile(sys.argv[1])

sheet_names = df_completo.sheet_names
salas = pd.read_excel(sys.argv[1], sheet_name=sheet_names[0])

df = pd.read_excel(sys.argv[1], sheet_name=sheet_names[1:])

jpter = pd.ExcelFile(sys.argv[2])
sheet_names_jpter = jpter.sheet_names
vagas = pd.read_excel(sys.argv[2], sheet_name=sheet_names_jpter[0:])

ingressantes = pd.read_excel(sys.argv[3])

espelho = pd.read_excel(sys.argv[4])

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

    colunas = vagas[sheet].columns

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
        
        if df[sheet].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
            index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])

            df[sheet].loc[d, 'Vagas por disciplina'] = 0
            for i in range(len(lista2)):
                df[sheet].loc[d, 'Vagas por disciplina'] += vagas[sheet].loc[index, colunas[lista2[i]]]
            
        else:
            df[sheet].loc[d, 'Vagas por disciplina'] = 0

        if pd.isna(df[sheet].loc[d, 'Observações']) or df[sheet].loc[d, 'Observações'] != 0:
            if "Ingressantes" in str(df[sheet].loc[d, 'Observações']):
                if df[sheet].loc[d, 'Disciplina (código)'] in ingressantes['Disciplina (código)'].tolist():
                    index = (ingressantes['Disciplina (código)'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
                    df[sheet].loc[d, 'Vagas por disciplina'] += ingressantes.loc[index, 'Ingressantes']

        if pd.isna(df[sheet].loc[d, 'Observações']) or df[sheet].loc[d, 'Observações'] != 0:
            if "Espelho" in str(df[sheet].loc[d, 'Observações']):
                if df[sheet].loc[d, 'Disciplina (código)'] in espelho['Disciplina (código)'].tolist():
                    index = (espelho['Disciplina (código)'].tolist()).index(df[sheet].loc[d, 'Disciplina (código)'])
                    df[sheet].loc[d, 'Vagas por disciplina'] += espelho.loc[index, 'Inscritos']
            

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
    
    for d in range(len(df[sheet_names[5]]['Disciplina (código)'])):
        
        if df[sheet_names[5]].loc[d, 'Disciplina (código)'] in vagas[sheet]['Disciplina'].tolist():
            index = (vagas[sheet]['Disciplina'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])

            df[sheet_names[5]].loc[d, 'Vagas por disciplina'] = 0
            for i in range(len(lista2)):
                df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += vagas[sheet].loc[index, colunas[lista2[i]]]
            
        elif pd.isna(df[sheet_names[5]].loc[d, 'Vagas por disciplina']):
            df[sheet_names[5]].loc[d, 'Vagas por disciplina'] = 0
            
for d in range(len(df[sheet_names[5]]['Disciplina (código)'])):            
    if pd.isna(df[sheet_names[5]].loc[d, 'Observações']) or df[sheet_names[5]].loc[d, 'Observações'] != 0:
        if "Ingressantes" in str(df[sheet_names[5]].loc[d, 'Observações']):
            if df[sheet_names[5]].loc[d, 'Disciplina (código)'] in ingressantes['Disciplina (código)'].tolist():
                index = (ingressantes['Disciplina (código)'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
                df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += ingressantes.loc[index, 'Ingressantes']
    if pd.isna(df[sheet_names[5]].loc[d, 'Observações']) or df[sheet_names[5]].loc[d, 'Observações'] != 0:
            if "Espelho" in str(df[sheet_names[5]].loc[d, 'Observações']):
                if df[sheet_names[5]].loc[d, 'Disciplina (código)'] in espelho['Disciplina (código)'].tolist():
                    index = (espelho['Disciplina (código)'].tolist()).index(df[sheet_names[5]].loc[d, 'Disciplina (código)'])
                    df[sheet_names[5]].loc[d, 'Vagas por disciplina'] += espelho.loc[index, 'Inscritos']
            


full_name = sys.argv[5]
file_path = os.path.join(os.getcwd(), "Saídas da Interface", "Planilhas de Dados", full_name)

    
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
    # Retorno o erro para o console
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)