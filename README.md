from tkinter import ttk
import customtkinter as ctk
import pandas as pd
from tkinter import filedialog
from tkinter import messagebox
import os
from time import sleep
import numpy as np
import logging
from datetime import datetime
from collections import defaultdict
from bs4 import BeautifulSoup
 
 
def update_window(operation):
    # Limpa o frame de conteúdo
    for widget in content_frame.winfo_children():
        widget.destroy()
   
    if operation == "Coleta de Protocolos":
        protocolos()
 
    elif operation == "Crítica e Conforto":
        critica_conforto()
 
    elif operation == "Correção de Centavos":
        correcao_1_centavo()
       
    elif operation == "Diferença no CDA":
        diferenca_cda()
    elif operation == "Pacotinhos":
        pacotinhos()
 
def protocolos():
 
    titulo = ctk.CTkLabel(content_frame, text="Batimento de Protocolo",font=("Arial", 20, "bold"))
    titulo.pack(pady=10)
 
    mes_label = ctk.CTkLabel(content_frame, text="Mês")
    mes_label.pack(pady=10)
 
    mes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    mes_combobox = ctk.CTkOptionMenu(content_frame, values=mes)
    mes_combobox.pack()
 
    ano_label = ctk.CTkLabel(content_frame, text="Ano")
    ano_label.pack()
 
    ano = ['2024','2025','2026','2027','2028','2029','2030']
    ano_combobox = ctk.CTkOptionMenu(content_frame, values=ano)
    ano_combobox.pack()
 
    data_label = ctk.CTkLabel(content_frame, text="Previsão de envio")
    data_label.pack(pady=10)
 
    data_entry = ctk.CTkEntry(content_frame)
    data_entry.pack()
    data_atual = datetime.now().strftime('%d/%m/%Y')
    data_entry.insert(0, data_atual)
 
 
    arquivo_label = ctk.CTkLabel(content_frame, text="Nenhum arquivo selecionado")
    arquivo_label.pack(side='bottom',pady=20)
 
    botao = ctk.CTkButton(content_frame, text="Iniciar", command=lambda: selecionar_arquivo(arquivo_label,mes_combobox,ano_combobox,switch_var,data_entry))
    botao.pack(side='bottom')
 
 
    switch_var = ctk.StringVar(value="CDA")
   
    cda_radiobutton = ctk.CTkRadioButton(content_frame, text="CDA", variable=switch_var, value="CDA")
    cda_radiobutton.pack(side='left', padx=50)
   
    balancete_radiobutton = ctk.CTkRadioButton(content_frame, text="Balancete", variable=switch_var, value="Balancete")
    balancete_radiobutton.pack(side='right', padx=50)
 
    def selecionar_arquivo(arquivo_label,mes_combobox,ano_combobox,switch_var,data_entry):
        meses = {'Janeiro': '01',
            'Fevereiro': '02',
            'Março': '03',
            'Abril': '04',
            'Maio': '05',
            'Junho': '06',
            'Julho': '07',
            'Agosto': '08',
            'Setembro': '09',
            'Outubro': '10',
            'Novembro': '11',
            'Dezembro': '12'}
       
        messagebox.showwarning("AVISO","O arquivo de protocolo não deve haver mesclagem e deverá possuir uma distância de 2 linhas entre cada protocolo.")
       
        def criar_batimento():  
            df1 = pd.read_excel(r"C:\Users\F9342792\OneDrive - Banco do Brasil S.A\General - CONTA FI\Controle FI\Controle_Espelho_FI.xlsx")   # Arquivo de base *ATUALIZAR*
            #df1['DRIVE'] = df1['DRIVE'].fillna(0).astype(int)
            df1['GFI'] = df1['GFI'].fillna(0).astype(int)
            df1 = df1[df1['SITUACAO'] == 'A']
            cabecalho = ['NOMES DOS FUNDOS','CNPJ','INFORME','STATUS','SITUAÇÃO','PROTOCOLO','COMPETÊNCIA','DATA PREVISTA PARA ENVIO','DATA DE ENVIO', 'MULTA', 'DRIVE', 'GFI']
            df2 = pd.DataFrame(columns=cabecalho)
            df2['NOMES DOS FUNDOS'] = df1['NOME FUNDOS - CVM']
            df2['CNPJ'] = df1['CNPJ']
            df2['DRIVE'] = df1['DRIVE']
            df2['GFI'] = df1['GFI']
            df2['STATUS'] = ['EM FUNCIONAMENTO NORMAL' if cnpj else '' for cnpj in df2['CNPJ']]
            return df2
 
        try:
 
            file_path = filedialog.askopenfilename(filetypes=[("Protocolo", "*.xlsx")])
            arquivo_label.configure(text=file_path)          
 
            # Lendo o segundo arquivo do Excel
            df2 = pd.read_excel(file_path, usecols=[0, 1])
           
            df1 = criar_batimento()
 
            # Criando um dicionário a partir de cada 15 linhas do segundo arquivo
            dataframes = []
            total_protocolos = 0
            localizados = []
            df2 = df2.dropna(how='all')
 
            for i in range(0, len(df2), 15):
                df = df2.iloc[i:i+15]
                dataframes.append(dict(zip(df.iloc[:, 0], df.iloc[:, 1])))
                total_protocolos +=1
            df.head(15)
           
            for index, row in df1.iterrows():
                try:
                    if pd.isna(row[0]):
                        continue
                    value =   row[0]
 
                    '''value = row[0]
                    print(value)'''
                   
                    for d in dataframes:
                        if value in d.values():
 
                            # Adiciona o dicionário à lista de dicionários localizados
                            localizados.append(d)
                            # Copiando o valor da chave "Status:"
                            content = d.get("Status:")
                            # Colando o conteúdo na quinta coluna do primeiro arquivo
                            df1.iloc[index, 4] = content
 
                            # Copiando o valor da chave "Informe:"
                            content = d.get("Informe:")
                            # Colando o conteúdo na terceira coluna do primeiro arquivo
                            df1.iloc[index, 2] = content
 
                            # Copiando o valor da chave "Nº Protocolo:"
                            content = d.get("Nº Protocolo:")
                            # Colando o conteúdo na sexta coluna do primeiro arquivo
                            df1.iloc[index, 5] = content
 
                            try:
                                # Copiando o valor da chave "Competência::"
                                content = d.get("Competência:")
                                # Convertendo a data para o formato desejado
                                formatted_date = content.strftime('%b/%y')
                                # Colando o conteúdo na sétima coluna do primeiro arquivo
                                df1.iloc[index, 6] = formatted_date
                            except Exception as e:
                                print(f"Ocorreu um erro: {e}")
 
                            # Copiando o valor da chave "Nº Protocolo:"
                            content = d.get("Data Ação:")
                            content = content[:10]
                            # Colando o conteúdo na nona coluna do primeiro arquivo
                            df1.iloc[index, 8] = content
 
                            # Colando o conteúdo na oitava coluna do primeiro arquivo
                            df1.iloc[index, 7] = data_entry.get()
                            if df1.iloc[index, 7]>= df1.iloc[index, 8]:
                                df1.iloc[index, 9] = "Não"
                            else:
                                df1.iloc[index, 9] = "Sim"
                except Exception as e:
                    print(f"Erro na linha {index + 2}: {e}")  # +2 para ajustar ao índice do Excel (1 para cabeçalho e 1 para índice base 0)
 
            df1 = df1.dropna(subset=['NOMES DOS FUNDOS'])
           
 
            fundo_perdido=[]
            # Verifica cada dicionário em dataframes
            print("Abaixo será exibido valores que estão nos protocolos e não estão no arquivo base: \n")
            for d in dataframes:
                # Se o dicionário não foi localizado
                if d not in localizados:
                    # Imprime o valor do "N° Protocolo:"
                    print(d.get("Participante:"))
                    fundo_perdido.append(d.get("Participante:"))
 
           
 
 
 
            mes = mes_combobox.get()
            mes_numero = meses[mes]
            ano = ano_combobox.get()
 
            if switch_var.get() == "CDA":
                nome_arquivo = f'BATIMENTO_CDA_{mes_numero}{ano}.xlsx'
                log_name = f'LOG_BATIMENTO_CDA_{mes_numero}{ano}.log'
            else:
                nome_arquivo = f'BATIMENTO_BALANCETES_{mes_numero}{ano}.xlsx'
                log_name = f'LOG_BATIMENTO_BALANCETE_{mes_numero}{ano}.log'
 
            # Define o caminho do arquivo na área de trabalho
            desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')
            '''desktop_path = fr'\\SRIRJO3A301\aplicbb\MER\Contratados\Estagiários\Marcus Vinicius\Saída Procolos'''
            caminho = os.path.join(desktop_path, nome_arquivo)
 
            # Cria o log
            logging.basicConfig(filename=os.path.join(desktop_path, f'{log_name}'), level=logging.INFO)
            logging.info(f'Abaixo serão exibidos os fundos que estão nos protocolos e não estão no arquivo base: \n')
            for fundo in fundo_perdido:
                logging.info(fundo)
 
            df1.to_excel(caminho, index=False)
            messagebox.showinfo("Sucesso",f"Finalizado, arquivos enviados para a {desktop_path}")
        except Exception as e:
            messagebox.showerror("Erro",{e})
            print(e)
 
def critica_conforto():
    import csv
    import os
    from tkinter import filedialog, messagebox
    import pandas as pd
    import shutil
 
    # Solicita ao usuário que selecione um diretório
    global directory, desktop_dir, file_to_move
    try:
 
        directory = filedialog.askdirectory()
 
    except NameError as e:
 
        print(f'Ocorreu um erro de acesso não permitido: {e}')
 
    desktop_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    desktop_dir = os.path.join(desktop_path, 'Sem natureza trocada')
    file_to_move = []
 
    def caso_duplo(directory):
        for file in os.listdir(directory):
            if file.endswith(".csv"):
                file_path = os.path.join(directory, file)
                df = pd.read_csv(file_path, delimiter=";", decimal=",", encoding='latin1')
                df.columns = df.iloc[-1]
                df = df[:-1]
                df = df.dropna(subset=['SldAtu'])
                df['SldAtu'] = df['SldAtu'].replace('[\+\-]', '', regex=True)
                df['SldAtu'] = df['SldAtu'].replace(',', '.', regex=True)
                df['SldAtu'] = df['SldAtu'].astype(str)
                df['SldAtu'] = df['SldAtu'].dropna()
                df['SldAtu'] = df['SldAtu'].astype(float).abs().round(2).astype(str)
                df['Nome'] = df['Nome'].str.strip()
               
 
                if not df['SldAtu'].empty and any(df['SldAtu'] == "0.0"):
                    # Verifica se 'Bloqueio' está na coluna 'Nome'
                    if 'Bloqueio' in df['Nome'].values:
                        # Armazena o valor da coluna 'SldAtu' na variável 'bloqueio_value'
                        bloqueio_value = df.loc[df['Nome'] == 'Bloqueio', 'SldAtu'].values[0]
                       
                        # Verifica se todos os valores na coluna 'SldAtu' são iguais
                        if all((df['SldAtu'] == bloqueio_value) | (df['SldAtu'] == "0.0")):
                            print(f'Bloqueio e 0.00: {file}')
                            file_to_move.append(file)
 
    def bloqueio(directory):
        for file in os.listdir(directory):
            if file.endswith(".csv"):
                file_path = os.path.join(directory, file)
                df = pd.read_csv(file_path, delimiter=";", decimal=",", encoding='latin1')
                if df.empty:
                    print(f'Arquivo vazio: {file}')
                    continue
                print(file)
                # Nomeia as colunas e remove a última linha
                df.columns = df.iloc[-1]
                df = df[:-1]
                df['Nome'] = df['Nome'].str.strip()
 
                # Verifica se 'Bloqueio' está na coluna 'Nome'
                if 'Bloqueio' in df['Nome'].values:
                    # Remove valores nulos da coluna 'SldAtu'
                    sld_atu_sem_nulos = df['SldAtu'].dropna()
                   
                    # Verifica se todos os valores na coluna 'SldAtu' são iguais
                    if sld_atu_sem_nulos.nunique() == 1:
                        print(f'Bloqueio: {file}')
                        file_to_move.append(file)
 
    def sem_trocada(directory):
        # Lista todos os arquivos CSV no diretório selecionado
        csv_files = [f for f in os.listdir(directory) if f.endswith('.csv')]
 
        # Verifica cada arquivo CSV
        for file in csv_files:
            file_path = os.path.join(directory, file)
            with open(file_path, newline='') as csvfile:
                reader = csv.reader(csvfile, delimiter=';')
                for row in reader:
                    # Verifica se a linha contém o texto "Não Existem Contas Com Natureza Trocada"
                    if any('Não Existem Contas Com Natureza Trocada' in cell for cell in row):
                        print(f'Sem Natureza trocada no arquivo: {file}')
                        file_to_move.append(file)
                        break
 
    def valor_nulo(directory):
        for file in os.listdir(directory):
            if file.endswith(".csv"):
                file_path = os.path.join(directory, file)
                df = pd.read_csv(file_path, delimiter=";", decimal=",", encoding='latin1')
                df.columns = df.iloc[-1]
                df = df[:-1]
                df = df.dropna(subset=['SldAtu'])
                df['SldAtu'] = df['SldAtu'].replace('[\+\-]', '', regex=True)
                df['SldAtu'] = df['SldAtu'].replace(',', '.', regex=True)
                df['SldAtu'] = df['SldAtu'].astype(str)
                df['SldAtu'] = df['SldAtu'].astype(float).abs().round(2).astype(str)
 
                if not df['SldAtu'].empty and all(df['SldAtu'] == "0.0"):
                    print(f"Valor 0,00 no arquivo:{file}")
                    # Copiar o arquivo para a pasta na área de trabalho
                    file_to_move.append(file)
 
 
    bloqueio(directory)
    sem_trocada(directory)
    valor_nulo(directory)
    caso_duplo(directory)
   
    informe = ctk.CTkLabel(content_frame, text= "Pasta verificada")
    informe.pack(side='bottom',pady=20)
 
    # Criar o diretório na área de trabalho se ele não existir
    if not os.path.exists(desktop_dir):
        os.makedirs(desktop_dir)
 
    # Copiar todos os arquivos na lista para a pasta na área de trabalho
    for file in file_to_move:
        src_file_path = os.path.join(directory, file)
        dst_file_path = os.path.join(desktop_dir, file)
        shutil.move(src_file_path, dst_file_path)
 
    messagebox.showinfo('Aviso', f'Arquivos copiados para {desktop_dir}')
 
    #R0059 é um exemplo de condição dupla
 
def correcao_1_centavo():
   
    titulo = ctk.CTkLabel(content_frame, text="Correção Balancete",font=("Arial", 20, "bold"))
    titulo.pack(pady=10)
   
    arquivo_label = ctk.CTkLabel(content_frame, text="Nenhum arquivo selecionado")
    arquivo_label.pack(side='bottom',pady=20)
 
    botao = ctk.CTkButton(content_frame, text="Iniciar", command=lambda: selecionar_arquivo(arquivo_label,switch_var))
    botao.pack(side='bottom')
 
 
    switch_var = ctk.StringVar(value="padrao")
   
    cda_radiobutton = ctk.CTkRadioButton(content_frame, text="Arquivo comum", variable=switch_var, value="padrao")
    cda_radiobutton.pack(side='left', padx=50)
   
    balancete_radiobutton = ctk.CTkRadioButton(content_frame, text="Arquivo Alternativo", variable=switch_var, value="alternativo")
    balancete_radiobutton.pack(side='right', padx=10)
   
    def selecionar_arquivo(arquivo_label,switch_var):
       
        os.system('cls' if os.name == 'nt' else 'clear') # Limpa o console
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos CVM", "*.CVM")])
 
        with open(caminho, 'r') as arquivo:
 
            nome, extensao = os.path.splitext(arquivo.name)
 
            ultima_linha = arquivo.readlines()[-1] # Lê a última linha do arquivo
            ultima_linha = ultima_linha.lstrip() # Remove espaços em branco
            ultima_linha =  ultima_linha.ljust(71)
 
            arquivo.seek(0, os.SEEK_SET) # Leva o ponteiro até o início do arquivo
            df = pd.read_csv(arquivo, delimiter=r'\s+', skipfooter=1)
            # Obtém os nomes das colunas
            colunas = df.columns
 
            # Converte os nomes das colunas em uma string separada por 4 espaços
            colunas_string = f"{colunas[0]}{'        '}{colunas[1]}"
 
            # Verifica se a string das colunas tem menos de 71 caracteres
            colunas_string = colunas_string.ljust(71)
 
            # O método iloc busca pela posição do índice. Neste caso, toda a linha da coluna 2
            lista = df.iloc[:, 1].str.extract(r'([+-].*)$', expand=False)
            df.iloc[:, 1] = df.iloc[:, 1].str.split('+').str[0]
            df.iloc[:, 1] = df.iloc[:, 1].str.strip('-').str.lstrip(' ')
            df.iloc[:, 1] = df.iloc[:, 1].str.split('-').str[0]
            df.iloc[:, 1] = df.iloc[:, 1].str.split('-').str[0]
 
 
            # Cria um dicionário com as 2 colunas do dataframe
            # O método "zip" é usado para criar pares de elementos das 2 colunas
            dicionario =dict(zip(df.iloc[:, 0],df.iloc[:, 1]))
            print(dicionario)  
 
            # Cria um novo dicionário, as chaves(k) continuam as mesmas
            # Os valores(v) são convertidos em inteiros
            dicionario_num = {k: int(v) for k, v in dicionario.items()}
           
            if switch_var.get() == "alternativo":
                if dicionario_num.get(13300003) is not None:
                    chave = 13000004 if dicionario_num.get(13000004) is not None else 13200004
                    sub = abs(dicionario_num[chave] - dicionario_num[13300003])
                    diferenca = abs(dicionario_num[30300000] - sub)
                    print('Sub: '+ str(sub))
                else:
                    chave = 13000004 if dicionario_num.get(13000004) is not None else 13200004
                    diferenca = abs(dicionario_num[chave] - dicionario_num[30300000])
 
                print('Diferenca: '+ str(diferenca))
 
                chave = 13000004 if dicionario_num.get(13000004) is not None else 13200004
                if dicionario_num[chave] > dicionario_num[30300000]:
                    dicionario_num[30330771] += diferenca
                    dicionario_num[30000001] += diferenca
                    dicionario_num[30300000] += diferenca
                    dicionario_num[30330001] += diferenca
                    dicionario_num[39999993] += diferenca
                    dicionario_num[90000003] += diferenca
                    dicionario_num[90300002] += diferenca
                    dicionario_num[90320006] += diferenca
                    dicionario_num[99999995] += diferenca
                else:
                    dicionario_num[30330771] -= diferenca
                    dicionario_num[30000001] -= diferenca
                    dicionario_num[30300000] -= diferenca
                    dicionario_num[30330001] -= diferenca
                    dicionario_num[39999993] -= diferenca
                    dicionario_num[90000003] -= diferenca
                    dicionario_num[90300002] -= diferenca
                    dicionario_num[90320006] -= diferenca
                    dicionario_num[99999995] -= diferenca
 
           
            # Verifica se existe a conta 1.3.1.15.00
            elif dicionario_num.get(13115009) is not None:
                cotas_fi = dicionario_num[13115009] + dicionario_num[13185709]          
                print(cotas_fi - dicionario_num[30330001])
                print(type(dicionario_num))
 
                cotas_fi = dicionario_num[13115009] + dicionario_num[13185709]
                print(cotas_fi)
 
                if cotas_fi > dicionario_num[30330771]:
                    diferenca = cotas_fi - dicionario_num[30330771]
                    dicionario_num[30330771] += diferenca
                    dicionario_num[30000001] += diferenca
                    dicionario_num[30300000] += diferenca
                    dicionario_num[30330001] += diferenca
                    dicionario_num[39999993] += diferenca
                    dicionario_num[90000003] += diferenca
                    dicionario_num[90300002] += diferenca
                    dicionario_num[90320006] += diferenca
                    dicionario_num[99999995] += diferenca
 
                else:
                    diferenca = dicionario_num[30330771] - cotas_fi
                    dicionario_num[30330771] -= diferenca
                    dicionario_num[30000001] -= diferenca
                    dicionario_num[30300000] -= diferenca
                    dicionario_num[30330001] -= diferenca
                    dicionario_num[39999993] -= diferenca
                    dicionario_num[90000003] -= diferenca
                    dicionario_num[90300002] -= diferenca
                    dicionario_num[90320006] -= diferenca
                    dicionario_num[99999995] -= diferenca
 
                print(diferenca)
                '''13185709'''
            else:    
               
                diferenca = dicionario_num[13185008] - dicionario_num[30330771]
                diferenca=abs(diferenca)
 
                if dicionario_num[13185008] > dicionario_num[30330771]:
                    dicionario_num[30330771] += diferenca
                    dicionario_num[30000001] += diferenca
                    dicionario_num[30300000] += diferenca
                    dicionario_num[30330001] += diferenca
                    dicionario_num[39999993] += diferenca
                    dicionario_num[90000003] += diferenca
                    dicionario_num[90300002] += diferenca
                    dicionario_num[90320006] += diferenca
                    dicionario_num[99999995] += diferenca
 
                elif dicionario_num[13185008] < dicionario_num[30330771]:
                    dicionario_num[30330771] -= diferenca
                    dicionario_num[30000001] -= diferenca
                    dicionario_num[30300000] -= diferenca
                    dicionario_num[30330001] -= diferenca
                    dicionario_num[39999993] -= diferenca
                    dicionario_num[90000003] -= diferenca
                    dicionario_num[90300002] -= diferenca
                    dicionario_num[90320006] -= diferenca
                    dicionario_num[99999995] -= diferenca
 
 
            # Cria um novo dataframe usando o dicionário
            novo_df = pd.DataFrame({df.columns[0]: list(dicionario_num.keys()), df.columns[1]: list(dicionario_num.values())})
            novo_df.iloc[:, 0] = '00' + novo_df.iloc[:, 0].astype(str)
            novo_df.iloc[:, 1] = novo_df.iloc[:, 1].astype(str).str.zfill(18)
            novo_df.iloc[:, 1] = novo_df.iloc[:, 1].astype(str)
               
            # Define o sufixo
            sufixo = '_CORRIGIDO'
 
            # Cria o novo nome do arquivo
            novo_nome = f"{nome}{sufixo}{extensao}"
 
            # Salva o DataFrame em um arquivo com 4 espaços entre as colunas
            with open(novo_nome, 'w') as f:
                # Verifica se a string das colunas tem menos de 71 caracteres
                if len(colunas_string) < 71:
                    # Adiciona espaços em branco ao final da string até que ela tenha 71 caracteres
                    colunas_string = colunas_string.ljust(71)
 
                # Itera sobre as linhas do DataFrame novo_df
                # Escreve os nomes das colunas no início do arquivo
                f.write(colunas_string + '\n')
                for i, (_, row) in enumerate(novo_df.iterrows()):
                    # Formata a linha com 4 espaços entre as colunas
                    string_formatada = f"{row[0]}{'    '}{row[1]}{lista[i]}"
 
                    if len(string_formatada) < 71:
                        # Adiciona espaços em branco ao final da linha até que ela tenha 71 caracteres
                        string_formatada = string_formatada.ljust(71)
                   
                    # Escreve a linha no arquivo
                    f.write(string_formatada + '\n')    
 
                # Escreve a última linha no arquivo    
                f.write(ultima_linha)
           
            messagebox.showinfo("Aviso", f"O arquivo foi ajustado")
            with open(novo_nome, 'r') as f:
                for i, linha in enumerate(f):
                    print(f"Linha {i+1}: {len(linha.strip())}")
 
def diferenca_cda():
   
    titulo = ctk.CTkLabel(content_frame, text="Diferença no CDA",font=("Arial", 20, "bold"))
    titulo.pack(pady=10)
   
    arquivo_label = ctk.CTkLabel(content_frame, text="Nenhum arquivo selecionado")
    arquivo_label.pack(side='bottom',pady=20)
 
    botao = ctk.CTkButton(content_frame, text="Iniciar", command=lambda: selecionar_arquivo(arquivo_label,switch_var,nome_arquivo))
    botao.pack(side='bottom',pady=20)    
 
   
    nome_arquivo = ctk.CTkEntry(content_frame,width=70,placeholder_text="Ex:1,2,3")
    nome_arquivo.pack(side='bottom')  
   
    label_entry = ctk.CTkLabel(content_frame, text="Qual lista o arquivo pertence?")
    label_entry.pack(side='bottom',pady=10)
 
    switch_var = ctk.StringVar(value="padrao")
   
    lista_radiobutton = ctk.CTkRadioButton(content_frame, text="Lista", variable=switch_var, value="Lista")
    lista_radiobutton.pack(side='left', padx=50)
   
    consolidado_radiobutton = ctk.CTkRadioButton(content_frame, text="Consolidado", variable=switch_var, value="Consolidado")
    consolidado_radiobutton.pack(side='right', padx=50)
   
   
    def selecionar_arquivo(cda,switch_var,nome_arquivo):
       
        if switch_var.get() == "Lista":
            sufixo = nome_arquivo.get()
           
            try:
                # Carregar o arquivo CSV selecionado pelo usuário
                cda = pd.read_csv(filedialog.askopenfilename(title='Selecione o arquivo CSV da CDA'), header=None, sep=';', encoding='ANSI')
                cda.columns = cda.iloc[-1]
                cda = cda.iloc[:-1]
                cda = cda[['Carteira', 'PlCrt', 'PlCalc', 'Titulo']]
                cda.fillna(0, inplace=True)
 
                # Substituir vírgulas por pontos e converter para float
                cda['PlCalc'] = cda['PlCalc'].str.replace(',', '.').astype(float)
                cda['PlCrt'] = cda['PlCrt'].str.replace(',', '.').astype(float)
 
                # Calcular a diferença
                cda['Diferença'] = cda['PlCalc'] - cda['PlCrt']
 
                # Inicializar o dicionário
                diferença = defaultdict(list)
 
                # Obter as linhas onde 'Diferença' não é 0
                rows = cda.loc[cda['Diferença'] != 0]
 
                # Adicionar valores de 'Carteira' e 'Título' ao dicionário
                for index, row in rows.iterrows():
                    diferença[row['Carteira']].append((row['Titulo'], row['Diferença']))
 
                # Remover a chave 0 se existir
                if 0 in diferença:
                    del diferença[0]
 
                # Escrever as diferenças no arquivo na pasta downloads do usuário
                path_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
 
                print(path_downloads)
                with open(os.path.join(path_downloads, f'Diferenças_CDA_Lista_{sufixo}.txt'), 'w') as f:
                    if diferença:
                        f.write('              A(s) seguinte(s) diferença(s) foi(foram) encontrada(s):\n')
                        f.write('-'*90 + '\n')
                        for carteira, titulos in diferença.items():
                            for titulo, diff in titulos:
                                # Substituir '.' por ',' na diferença
                                diff_str = f'{diff:.2f}'.replace('.', ',')
                                titulo_str = f'Título: {titulo} Diferença: {diff_str}'
                                num_spaces = 88 - len(f'Carteira: {carteira} {titulo_str}')
                                f.write(f'|Carteira: {carteira} {titulo_str}{" " * num_spaces}|\n')
                                f.write('-'*90 + '\n')
 
 
                    else:
                        f.write('Nenhuma diferença encontrada\n')
                    messagebox.showinfo("Aviso", f"Arquivo finalizado e enviado para a pasta de downloads")
            except Exception as e:
                messagebox.showerror('Erro', f'Ocorreu um erro: {e}')
       
        else:
            try:
                defeito=None
                # Selecionar a pasta com os arquivos CSV da CDA
                pasta = filedialog.askdirectory(title='Selecione a pasta com os arquivos CSV da CDA')
                # Inicializar o dicionário para armazenar todas as diferenças
                todas_diferencas = defaultdict(list)
               
                # Iterar sobre todos os arquivos CSV na pasta selecionada
                for arquivo in os.listdir(pasta):
                    if arquivo.endswith('.csv'):
                        caminho_arquivo = os.path.join(pasta, arquivo)
                       
                        # Carregar o arquivo CSV
                        cda = pd.read_csv(caminho_arquivo, header=None, sep=';', encoding='ANSI')
                        cda.columns = cda.iloc[-1]
                        cda = cda.iloc[:-1]
                        cda = cda[['NomeCrt' ,'PlCrt', 'PlCalc', 'Titulo']]
                        cda.fillna(0, inplace=True)
                       
                        # Substituir vírgulas por pontos e converter para float
                        cda['PlCalc'] = cda['PlCalc'].str.replace(',', '.').astype(float)
                        cda['PlCrt'] = cda['PlCrt'].str.replace(',', '.').astype(float)
                       
                        # Calcular a diferença
                        cda['Diferença'] = cda['PlCalc'] - cda['PlCrt']
                        if cda['Diferença'].any() != 0:
                            print(f'Arquivo: {arquivo}')
                            defeito=1
 
                           
                       
                        # Obter as linhas onde 'Diferença' não é 0
                        rows = cda.loc[cda['Diferença'] != 0]
                       
                       
                       
                        # Adicionar valores de 'Carteira' e 'Título' ao dicionário
                        for index, row in rows.iterrows():
                            todas_diferencas[row['NomeCrt']].append((row['Titulo'], row['Diferença']))
               
                # Remover a chave 0 se existir
                if 0 in todas_diferencas:
                    del todas_diferencas[0]
               
                # Escrever as diferenças no arquivo na pasta downloads do usuário
                path_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
               
                with open(os.path.join(path_downloads, 'Diferenças_CDA_Consolidados.txt'), 'w') as f:
                    if todas_diferencas:
                        f.write('              A(s) seguinte(s) diferença(s) foi(foram) encontrada(s):\n')
                        f.write('-'*90 + '\n')
                        for carteira, titulos in todas_diferencas.items():
                            for titulo, diff in titulos:
                                # Substituir '.' por ',' na diferença
                                diff_str = f'{diff:.2f}'.replace('.', ',')
                                titulo_str = f'Título: {titulo} Diferença: {diff_str}'
                                num_spaces = 88 - len(f'Carteira: {carteira} {titulo_str}')
                                f.write(f'|Carteira: {carteira} {titulo_str}{" " * num_spaces}|\n')
                                f.write('-'*90 + '\n')
                    else:
                        if defeito==1:
                            f.write('Há diferenças no arquivo\n')
                            print(f'Arquivo com defeito: {arquivo}')
                        f.write('Nenhuma diferença encontrada\n')
                    messagebox.showinfo("Aviso", f"Arquivo finalizado e enviado para a pasta de downloads")
            except Exception as e:
                messagebox.showerror('Erro', f'Ocorreu um erro: {e}')
 
def pacotinhos():
    toplevel_window = None  # Declaração inicial da variável
 
    def ajuda():
        nonlocal toplevel_window  # Declaração de variável não local
 
        def create_textbox_with_optional_image(parent, text, image_path=None):
            caixa_texto = ctk.CTkTextbox(parent, font=("Arial", 20))
            caixa_texto.pack(expand=True, fill='both', padx=10, pady=10)
            caixa_texto.insert('1.0', text)
                   
        def posicionar_esquerda(janela, largura, altura):
            pos_x = 300  # Posição X fixa à esquerda
            pos_y = 400  # Posição Y fixa
            janela.geometry(f'{largura}x{altura}+{pos_x}+{pos_y}')
 
        if toplevel_window is None or not toplevel_window.winfo_exists():
 
            toplevel_window = ctk.CTkToplevel()
            posicionar_esquerda(toplevel_window, 850, 350)
            toplevel_window.title("Ajuda")
            toplevel_window.lift()
            toplevel_window.focus_force()
           
            # Cria abas no top level
            abas = ctk.CTkTabview(toplevel_window)
            abas.pack(expand=True, fill='both', padx=10, pady=10)
           
            # Adiciona uma aba chamada "Informações"
            aba_info = abas.add("Informações")
            # Adiciona a caixa de texto dentro da aba "Informações"
            create_textbox_with_optional_image(aba_info, "Este programa foi desenvolvido para extrair informações de arquivos HTML e Excel.\n\n")
           
            aba_atualizar = abas.add("Como Coletar o HTML")
            create_textbox_with_optional_image(aba_atualizar, "Teste ")
        else:
            toplevel_window.lift()
            toplevel_window.focus_force()
           
    def selecionar_arquivo():
        global arquivo_selecionado
        arquivo = filedialog.askopenfile(title="Selecione o arquivo HTML", filetypes=(("HTML files", "*.html"), ("All files", "*.*")))
        if arquivo:
            arquivo_selecionado = arquivo.name
            label_arquivo.configure(text="Arquivo HTML selecionado", text_color="yellow")
   
    titulo = ctk.CTkLabel(content_frame, text="Pacotinhos", font=("Arial", 20, "bold"))
    titulo.pack(pady=10)
 
    botao_ajuda = ctk.CTkButton(content_frame, text="Ajuda", command=ajuda)
    botao_ajuda.pack(pady=20)
    label_arquivo = ctk.CTkLabel(content_frame, text="Nenhum arquivo selecionado")
    label_arquivo.pack(side='bottom', pady=20)
 
    tabview = ctk.CTkTabview(content_frame)
    tabview.pack(expand=True, fill='both', padx=10, pady=10)
 
    aba_1 = tabview.add("Mensal")
    aba_2 = tabview.add("Anual")
   
    # Adiciona o botão a ambas as abas
    label_aba1 = ctk.CTkLabel(aba_1, text="Selecione o mês e o arquivo HTML",font=("Arial", 15, "bold"))
    label_aba1.pack(pady=10)
    # Cria um botão dentro da aba "Mês"
    meses=['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ']
    mes_combobox = ctk.CTkComboBox(aba_1, values=meses)
    mes_combobox.pack(pady=10)
    botao_arquivo_aba_1 = ctk.CTkButton(aba_1, text="Selecionar Arquivo", command=selecionar_arquivo)
    botao_arquivo_aba_1.pack(pady=20)
 
    label_aba2 = ctk.CTkLabel(aba_2, text="Selecione o arquivo HTML",font=("Arial", 15, "bold"))
    label_aba2.pack(pady=34)
    botao_arquivo_aba_2 = ctk.CTkButton(aba_2, text="Selecionar Arquivo", command=selecionar_arquivo)
    botao_arquivo_aba_2.pack(pady=20)
 
    botao_iniciar_codigo_aba_1 = ctk.CTkButton(aba_1, text="Iniciar", command=lambda: extracao_vermelhos('Mês', mes_combobox.get()))
    botao_iniciar_codigo_aba_1.pack(pady=20)
 
    botao_iniciar_codigo_aba_2 = ctk.CTkButton(aba_2, text="Iniciar", command=lambda: extracao_vermelhos('Anual'))
    botao_iniciar_codigo_aba_2.pack(pady=20)
 
    switch_var = ctk.StringVar(value="Azul")
   
    cda_radiobutton = ctk.CTkRadioButton(aba_1, text="Azul", variable=switch_var, value="Azul")
    cda_radiobutton.pack(side='left', padx=50)
   
    balancete_radiobutton = ctk.CTkRadioButton(aba_1, text="Vermelho", variable=switch_var, value="Vermelho")
    balancete_radiobutton.pack(side='right', padx=50)
 
 
    def extracao_vermelhos(mes_ou_ano,selecao_mes=None):
        # Variável de escolha: "Mensal" ou "Anual"
        escolha = mes_ou_ano
 
        # Carregar o conteúdo do arquivo HTML
        with open(arquivo_selecionado, 'r', encoding='utf-8') as file:
            html_content = file.read()
 
        # Parseando o HTML
        soup = BeautifulSoup(html_content, 'html.parser')
 
        # Carregar o arquivo Excel
        excel_file = r"C:\Users\F9342792\OneDrive - Banco do Brasil S.A\General - CONTA FI\Controle FI\Controle_Espelho_FI.xlsx"
        df = pd.read_excel(excel_file,sheet_name="Pacotinho_FI")
        # Filtrar a coluna 'SITUACAO' para valores "A"
        filtered_df = df[df['SITUACAO'] == 'A']
 
        # Lista para armazenar os resultados
        results = []
 
        # Encontrando todas as tags <td> principais
        main_tds = soup.find_all('td')
 
        # Iterando sobre cada tag <td> principal
        for main_td in main_tds:
            # Encontrando a tabela interna
            inner_table = main_td.find('table', bgcolor="#FFFFF0")
           
            if inner_table:
                # Extraindo os meses com a cor vermelha
                red_months = []
 
                if switch_var.get() == "Azul":
 
                    cor_pesquisa = "#0000AA"
                else:
                    cor_pesquisa = "#FF0000"
 
                # Cor em Azul :"#0000AA"   #######################################################################################################################
                # Cor em vermelho : "#FF0000"
 
                for font_tag in inner_table.find_all('font', color=cor_pesquisa):
 
                    if escolha == "Mês":
                        if font_tag.text == selecao_mes:
                            red_months.append(font_tag.text)
                    else:  # Anual
                        red_months.append(font_tag.text)
               
                # Verificando se há meses em vermelho
                if red_months:
                    item_name = main_td.find('b').text.strip()
                   
                    # Comparar com os nomes do Excel
                    if item_name in filtered_df['NOME FUNDOS - CVM'].values:
                        results.append([item_name, ', '.join(red_months)])
 
        # Criar um DataFrame com os resultados
        if cor_pesquisa == "#FF0000":
            results_df = pd.DataFrame(results, columns=['Nome do Fundo', 'Meses em Vermelho'])
            nome_saida ='resultados_fundos_vermelho'
        else:
            results_df = pd.DataFrame(results, columns=['Nome do Fundo', 'Meses em Azul'])
            nome_saida ='resultados_fundos_azul'
 
        caminho_download = os.path.expanduser("~\Downloads")
 
        # Salvar o DataFrame em um arquivo Excel
        output_file = rf"{caminho_download}\{nome_saida}.xlsx"
        results_df.to_excel(output_file, index=False)
 
        print(f"Resultados salvos em {output_file}")
             
def create_button(frame, text):
    button = ctk.CTkButton(frame, text=text, command=lambda: update_window(text),fg_color="#0038A8",hover_color="#7B68EE",font=ctk.CTkFont(size=14,family="Arial"))
    button.pack(side='top', pady=10,fill='none',padx=5)
 
def change_appearance_mode_event(tema: str):
    ctk.set_appearance_mode(tema)
 
    if tema == "Light":
        logo_label.configure(text_color="#000000")  # Muda a cor do texto para preto
        appearance_mode_label.configure(text_color="#000000")  # Muda a cor do texto para preto
 
 
 
    elif tema == "Dark":
        logo_label.configure(text_color="#FEDD00")  # Muda a cor do texto para preto
        appearance_mode_label.configure(text_color="#FEDD00")  # Muda a cor do texto para preto
 
window = ctk.CTk()
window.geometry("600x600")
window.title("Projeto CVM")
 
 
frame_lateral = ctk.CTkFrame(window)
frame_lateral.pack(side='left', fill='y')
 
# Adiciona um separador vertical
separator = ttk.Separator(window,orient='vertical')
separator.pack(side='left', fill='y',padx=5)
 
content_frame = ctk.CTkFrame(window, width=100, height=100)
content_frame.pack(side='right', fill='both', expand=True)
 
logo_label = ctk.CTkLabel(frame_lateral, text="Menu Inicial", font=ctk.CTkFont(size=26, weight="bold"),text_color="#FEDD00")
logo_label.pack(pady=10)
 
#Cria novos botões
create_button(frame_lateral, "Crítica e Conforto")
create_button(frame_lateral, "Diferença no CDA")
create_button(frame_lateral, "Correção de Centavos")
create_button(frame_lateral, "Coleta de Protocolos")
create_button(frame_lateral, "Pacotinhos")
 
# Cria um CTkOptionMenu na base do frame_lateral
appearance_mode_label = ctk.CTkLabel(frame_lateral, text="Modo:",anchor = "s",text_color="#FEDD00")
appearance_mode_label.pack()
 
appearance_mode_optionemenu = ctk.CTkOptionMenu(frame_lateral, values=["Light", "Dark", "System"],
                                                     command=change_appearance_mode_event,fg_color="#0038A8")
appearance_mode_optionemenu.set("Dark")
appearance_mode_optionemenu.pack( fill='none',pady = 30)
 
sair = ctk.CTkButton(frame_lateral, text="Sair", command=window.quit,fg_color="#0038A8",hover_color="#7B68EE",font=ctk.CTkFont(size=14,family="Arial"))
sair.pack()
 
 
window.mainloop()
