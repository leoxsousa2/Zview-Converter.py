import importlib
import subprocess
# Atualiza o pip
subprocess.run(['pip', 'install', '--upgrade', 'pip'])
# lista de bibliotecas que serão verificadas e instaladas caso necessário
bibliotecas = ['importlib', 'subprocess', 'os', 'time', 'pandas', 'openpyxl']
# percorre a lista de bibliotecas
for biblioteca in bibliotecas:
    # verifica se a biblioteca está instalada
    try:
        importlib.import_module(biblioteca)
        #print(biblioteca + " já está instalada!")
    # caso não esteja instalada, utiliza o pip para fazer o download e instalação
    except ImportError:
        print("Instalando " + biblioteca + "...")
        subprocess.run(['pip', 'install', biblioteca])
    # importa a biblioteca após a instalação (ou se já estiver instalada)
    finally:
        globals()[biblioteca] = importlib.import_module(biblioteca)





import os
import time
import pandas as pd
import openpyxl  #importante fazer pip install openpyxl disso

os.system('cls' if os.name == 'nt' else 'clear')                    # limpa a tela do terminal

print("    ")
print('-->> Esse programa precisa do programa 7-Zip File Manager instalado!')
print("    ")
time.sleep(5)
os.system('cls' if os.name == 'nt' else 'clear')                    # limpa a tela do terminal
print("    ")
print('-->> Programa iniciado!!! Pode ir tomar um café.')
print("(Dev: Leo Sousa) vers: 1.0")
print("    ")
time.sleep(5)

def listar_arquivos_sem_extensao():
    caminho_da_pasta = os.path.abspath(os.path.dirname(__file__))
    extensao_desejada = ".rar"
    arquivos = [arquivo.split(".")[0] for arquivo in os.listdir(caminho_da_pasta) if arquivo.endswith(extensao_desejada)]
    if arquivos:
        for arquivo in arquivos:
            result = arquivo
    return result

def descompactarPasta():
    # Caminho para o executável do 7-Zip
    caminho_7z = r"C:\Program Files\7-Zip\7z.exe"
    # Caminho para o arquivo rar
    caminho_arquivo_rar = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + ".rar"
    # Caminho para a pasta de destino
    caminho_destino = os.path.abspath(os.path.dirname(__file__)) + "\\" +  listar_arquivos_sem_extensao()
    # Comando a ser executado
    comando = f'"{caminho_7z}" x "{caminho_arquivo_rar}" -o"{caminho_destino}" "-r"'
    # Executar o comando usando subprocess
    subprocess.run(comando, shell=True, check=True)
    time.sleep(1)
    os.system('cls' if os.name == 'nt' else 'clear')                    # limpa a tela do terminal
    print("    ")
    print('-->> Programa iniciado!!! Pode ir tomar um café.')
    print("    ")
    print("    ")
    print("Pasta descompactada")
    print("    ")

def renomear_pasta():
    nomeAntigo = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao()
    novoNome = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Original"
    os.rename(nomeAntigo, novoNome)
    print("    ")
    print("Criado a pasta Original")
    print("    ")

def criar_pastas():
    pasta1 = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Virgula"
    print("    ")
    print("Criado a pasta Virgula")
    print("    ")
    pasta2 = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Zview"
    print("    ")
    print("Criado a pasta Zview")
    print("    ")
    os.makedirs(pasta1)
    os.makedirs(pasta2)

def fazer_copia_da_pasta():
    print("    ")
    print('Fazendo cópia da pasta Original para  a pasta Vírgula')
    print("    ")
    time.sleep(1)
    caminho_origem = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Original"
    caminho_destino = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Virgula"
    # Comando para copiar a pasta usando xcopy no Windows
    comando = f'xcopy "{caminho_origem}" "{caminho_destino}" /E /I /Y'
    # Executa o comando usando o subprocess
    subprocess.run(comando, shell=True, check=True)
    time.sleep(1)
    os.system('cls' if os.name == 'nt' else 'clear')                    # limpa a tela do terminal
    print("    ")
    print('-->> Programa iniciado!!! Pode ir tomar um café.')
    print("    ")
    print("    ")
    print("Pasta descompactada")
    print("    ")
    print("    ")
    print("Criado a pasta Original")
    print("    ")
    print("    ")
    print("Criado a pasta Virgula")
    print("    ")
    print("    ")
    print("Criado a pasta Zview")
    print("    ")
    print("    ")
    print('Feito cópia da pasta Original para  a pasta Vírgula')
    print("    ")

def substituir_ponto_por_virgula_em_txt():
    caminho_pasta = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Virgula"
    # Percorre todos os arquivos na pasta
    for nome_arquivo in os.listdir(caminho_pasta):
        caminho_arquivo = os.path.join(caminho_pasta, nome_arquivo)
        # Verifica se o arquivo é um arquivo de texto (.txt)
        if nome_arquivo.lower().endswith('.txt') and os.path.isfile(caminho_arquivo):
            # Abre o arquivo, lê seu conteúdo, substitui pontos por vírgulas e salva de volta
            with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo:
                conteudo = arquivo.read()
                conteudo_modificado = conteudo.replace('.', ',')        
            with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo_modificado:
                arquivo_modificado.write(conteudo_modificado)
    print("    ")
    print("Substituindo o ponto por virgula de todos os arquivos da pasta Vírgula")
    print("    ")

def converter_txt_to_excel():
    pasta_origem = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Virgula"
    pasta_destino1 = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Zview"
    pasta_destino2 = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Virgula"
    # Itera sobre todos os arquivos na pasta de origem
    for arquivo_txt in os.listdir(pasta_origem):
        if arquivo_txt.endswith('.txt'):
            caminho_txt = os.path.join(pasta_origem, arquivo_txt)
            # Lê o arquivo .txt usando pandas
            df = pd.read_csv(caminho_txt, delimiter='\t')  # Ajuste o delimitador conforme necessário
            # Cria o nome do arquivo Excel na pasta de destino
            arquivo_excel = os.path.join(pasta_destino1, os.path.splitext(arquivo_txt)[0] + '.xlsx')
            # Salva o DataFrame em um arquivo Excel
            df.to_excel(arquivo_excel, index=False)
            arquivo_excel = os.path.join(pasta_destino2, os.path.splitext(arquivo_txt)[0] + '.xlsx')
            # Salva o DataFrame em um arquivo Excel
            df.to_excel(arquivo_excel, index=False)
    print("    ")
    print("Convertendo arquivos para excel e salvando nas pasta Virgula e na pasta Vview")
    print("    ")

def fazendo_ajustes_nos_arquivosEXCEL_pastaVirgula():
    print("    ")
    print("Copiando terceira coluna para a sétima - trocando sinal - na pasta Vírgula")
    print("    ")
    pasta_origem = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Virgula"
    pasta_destino = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Virgula"
    # Itera sobre todos os arquivos na pasta de origem
    for arquivo_excel in os.listdir(pasta_origem):
        if arquivo_excel.endswith('.xlsx'):
            caminho_excel = os.path.join(pasta_origem, arquivo_excel)
            # Lê o arquivo Excel usando pandas
            df = pd.read_excel(caminho_excel)
            # Copia a terceira coluna para a sétima coluna
            if len(df.columns) >= 3:
                df.insert(6, 'Z" (Ω)', df.iloc[:, 2])
                # Adiciona o sinal negativo à sétima coluna
                df.iloc[:, 6] = df.iloc[:, 6].apply(lambda x: f'-{x}')
                #df = df.drop(df.columns[2], axis=1) #excluindo a terceira coluna. Os codigo seguintes repetem até excluir as colunas que eu quero sempre excluindo a terceira coluna
                #df = df.drop(df.columns[2], axis=1)
                #df = df.drop(df.columns[2], axis=1)
                #df = df.drop(df.columns[2], axis=1)
                # Salva o DataFrame de volta no arquivo Excel na pasta de destino
                caminho_destino = os.path.join(pasta_destino, arquivo_excel)
                df.to_excel(caminho_destino, index=False)
                
def fazendo_ajustes_nos_arquivosEXCEL_pastaZview():
    print("    ")
    print("Copiando terceira coluna para a sétima - trocando sinal - excluindo colunas 3, 4, 5, 6 e 7 na pasta Zview")
    print("    ")
    pasta_origem = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Zview"
    pasta_destino = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Zview"
    # Itera sobre todos os arquivos na pasta de origem
    for arquivo_excel in os.listdir(pasta_origem):
        if arquivo_excel.endswith('.xlsx'):
            caminho_excel = os.path.join(pasta_origem, arquivo_excel)
            # Lê o arquivo Excel usando pandas
            df = pd.read_excel(caminho_excel)
            # Copia a terceira coluna para a sétima coluna
            if len(df.columns) >= 3:
                df.insert(6, 'Z" (Ω)', df.iloc[:, 2])
                # Adiciona o sinal negativo à sétima coluna
                df.iloc[:, 6] = df.iloc[:, 6].apply(lambda x: f'-{x}')
                df = df.drop(df.columns[2], axis=1) #excluindo a terceira coluna. Os codigo seguintes repetem até excluir as colunas que eu quero sempre excluindo a terceira coluna
                df = df.drop(df.columns[2], axis=1)
                df = df.drop(df.columns[2], axis=1)
                df = df.drop(df.columns[2], axis=1)
                #Salva o DataFrame de volta no arquivo Excel na pasta de destino
                caminho_destino = os.path.join(pasta_destino, arquivo_excel)
                df.to_excel(caminho_destino, index=False)

def excluindo_primeira_linha_excel_pastaZview():
    # Substitua 'caminho_da_pasta' pelo caminho da pasta que contém os arquivos Excel
    caminho_da_pasta = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Zview"
    # Lista todos os arquivos na pasta
    arquivos_excel = [arquivo for arquivo in os.listdir(caminho_da_pasta) if arquivo.endswith('.xlsx')]
    # Itera sobre cada arquivo Excel
    for arquivo in arquivos_excel:
        # Constrói o caminho completo do arquivo
        caminho_completo = os.path.join(caminho_da_pasta, arquivo)
        # Carrega o arquivo Excel em um DataFrame
        df = pd.read_excel(caminho_completo)
      # Carrega o arquivo Excel em um DataFrame, pulando a primeira linha (linha de título)
        df = pd.read_excel(caminho_completo, skiprows=[0])
        # Salva o DataFrame modificado de volta no arquivo Excel
        df.to_excel(caminho_completo, index=False)
    print("    ")
    print("Excluindo a primeira linha dos arquivos Excel")
    print("    ")

def converterExcel_para_TXT_e_excluirExcel():
    caminho_da_pasta = os.path.abspath(os.path.dirname(__file__)) + "\\" + listar_arquivos_sem_extensao() + "\\" + listar_arquivos_sem_extensao() + " Zview"

    # Lista todos os arquivos na pasta
    arquivos_excel = [arquivo for arquivo in os.listdir(caminho_da_pasta) if arquivo.endswith('.xlsx')]

    for arquivo in arquivos_excel:
        caminho_completo_excel = os.path.join(caminho_da_pasta, arquivo)
        # Carrega o arquivo Excel em um DataFrame
        df = pd.read_excel(caminho_completo_excel)
        # Substitui ponto por vírgula em todas as colunas do DataFrame
        df = df.map(lambda x: str(x).replace(',', '.'))
        # Gera o caminho para o arquivo de texto
        caminho_completo_txt = os.path.join(caminho_da_pasta, f"{os.path.splitext(arquivo)[0]}.txt")
        # Salva o DataFrame como arquivo de texto (tab-separated)
        df.to_csv(caminho_completo_txt, sep='\t', index=False)
        # Exclui o arquivo Excel após a conversão
        os.remove(caminho_completo_excel)

    print("    ")
    print("Convertendo para TXT e trocando ponto por vírgula na pasta Zview")
    print("    ")



time.sleep(1)        
descompactarPasta()

renomear_pasta()

time.sleep(1) 
criar_pastas()

time.sleep(1) 
fazer_copia_da_pasta()

time.sleep(1) 
substituir_ponto_por_virgula_em_txt()

time.sleep(1)
converter_txt_to_excel()

time.sleep(1)
converter_txt_to_excel()

time.sleep(1)
fazendo_ajustes_nos_arquivosEXCEL_pastaVirgula()

time.sleep(1)
fazendo_ajustes_nos_arquivosEXCEL_pastaZview()

time.sleep(1)
excluindo_primeira_linha_excel_pastaZview()

time.sleep(1)
converterExcel_para_TXT_e_excluirExcel()




time.sleep(1) 
print("    ")
print("Tudo pronto!!!")
print("    ")
time.sleep(1)

verificar = input("")