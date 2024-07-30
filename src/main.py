import pandas as pd
import os
import glob
import xlsxwriter
import pprint

# Caminho para ler os arquivos
folder_path = r'src\data\raw'

# Listar todos os arquivos de excel
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

# print(excel_files)

if not excel_files:
    print('Nenhum arquivo Excel encontrado.')
    exit()
else:
    
    # dataFrame - tabela na memoria para quardar dados dos arquivos 
    dfs = []
    
    for excel_file in excel_files:
        
        try:
            # EXTRAINDO DADOS
            # ler qrquivo do excel
            df_temp = pd.read_excel(excel_file)
            
            # pegar name de arquivos
            file_name = os.path.basename(excel_file)
            
            df_temp['filename'] = file_name
            
            # TRATANDO DADOS
            
            # Criando um nova coluna chamada location
            if 'brasil' in file_name.lower():
                df_temp['location'] = 'br'
            elif 'france' in file_name.lower():
                df_temp['location'] = 'fr'
            elif 'italian' in file_name.lower():
                df_temp['location'] = 'it'
                
            # Criando uma nova coluna chamada campaign
            df_temp['campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')
            
            # Guardando dados tratados dentro de uma dataframe comun
            dfs.append(df_temp)
            
        except Exception as e:
            print(f'Erro ao ler o aquivos {e}')
        
if dfs:
    # lODING(CARREGANDO DADOS)
    
    # concatena todas as tabelas salvas no dfs em uma unica tabela
    result = pd.concat(dfs, ignore_index=True)
    
    # Caminho de saida
    output_file = os.path.join('src', 'data', 'ready', 'clean.xlsx')
    
    # configurou o motor de escrita
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    
    # leva os dados do resultado a serem escritos no motor excel configurado
    result.to_excel(writer, index=False)
    
    # salva o arquivo de excel
    writer._save()
    
else:
    print('Nenhum arquivo Excel foi processado.')