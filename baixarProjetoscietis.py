import re
import os
import time
import shutil
import unicodedata
import pandas as pd 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options



# Configurações
arquivo_excel = 'artigos_CIETIS.xlsx'
download_dir = os.path.join(os.getcwd(), 'Downloads')
pasta_destino = os.path.join(os.getcwd(), 'Projetos Ciets')
tempo_espera = 10 #s
# modo_teste = True # Só processa a primeira linha pra teste
linhas_teste = None    # Use: 1 para uma linha, 5 para cinco linhas, None para todas


# função para limpar e limitar nome
def limpar_nome(texto, limite=250):
    texto = str(texto).strip()

    # Remove acentos
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')

    # Remove caracteres inválidos para arquivos
    texto = re.sub(r'[\\/*?:"<>|]', '', texto)
    texto = texto.replace('\n', ' ').replace('\r', '')

    # Limita o tamanho
    if len(texto) > limite:
        texto = texto[:limite].rstrip()

    return texto


# def extrair_nome_e_sobrenome(nome_completo):
#     partes = str(nome_completo).strip().split()
#     if len(partes) >= 2:
#         return f"{partes[0]} {partes[-1]}"
#     return partes[0]  # caso tenha só um nome






# Verificar se a planilha existe
if not os.path.exists(arquivo_excel):
    print(f"Arquivo '{arquivo_excel}' não encontrado na pasta: {os.getcwd()}")
    exit()
else:
    print(f"Planilha encontrada: {arquivo_excel}")



# Lê todas as planilhas do arquivo
xls = pd.ExcelFile(arquivo_excel)
planilhas = xls.sheet_names
print(f"📄 Encontradas as seguintes abas: {planilhas}")



# configurações do chrome
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    'download.default_directory': download_dir,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True
})
driver = webdriver.Chrome(options=chrome_options)



# Função para esperar o download
def aguardar_download(pasta, timeout=30):

    tempo_inicio = time.time()
    while True:
        arquivos = [f for f in os.listdir(pasta) if not f.endswith('.crdownload')]
        if arquivos:
            caminho = max([os.path.join(pasta, f) for f in arquivos], key=os.path.getctime)
            return caminho
        if time.time() - tempo_inicio > timeout:
            raise Exception("Tempo excedido esperando download.")
        time.sleep(2)




# Executando
xls = pd.ExcelFile(arquivo_excel)
planilhas = xls.sheet_names
print(f"Encontradas as seguintes planilhas: {planilhas}")


for nome_aba in planilhas:
    print(f"\n Processando aba: {nome_aba}")
    df = pd.read_excel(xls, sheet_name=nome_aba)


for i, linha in df.iterrows():
    if linhas_teste and i >= linhas_teste:
        break # roda até numero definido

    nome_autor = (linha[0])     
    autor = limpar_nome(nome_autor, limite=100) # Coluna A
    titulo = limpar_nome(str(linha[1]), limite=100)   # Coluna B
    link = str(linha[5])                              # Coluna F

    if not link.startswith("http"):
        print(f"[{i}] Link inválido: {link}")
        continue
        

    print(f"\n Testando projeto da linha {i+1}")
    print(f"AUTOR : {autor}")
    print(f"TÍTULO: {titulo}")
    print(f"LINK  : {link}")

    driver.get(link)
    time.sleep(3)

    print(" Aguardando download...")
    try:
        arquivo_baixado = aguardar_download(download_dir, tempo_espera)
    except Exception as e:
        print(f"Erro ao baixar: {e}")
        continue

    nova_extensao = os.path.splitext(arquivo_baixado)[-1]
    novo_nome = f"{autor}{nova_extensao}"
    
    subpasta_destino = str(linha.iloc[6]).strip()


    if not subpasta_destino or subpasta_destino.lower() == 'nan':
        print(f"[{i}] Caminho da subpasta vazio. Pulando...")
        continue


    # Garante que a pasta existe
    os.makedirs(subpasta_destino, exist_ok=True)

    caminho_destino = os.path.join(subpasta_destino, novo_nome)

    shutil.move(arquivo_baixado, caminho_destino)
    print(f"Arquivo salvo como: {novo_nome}")

    # break  # só a primeira linha

driver.quit()






# # ========== LOOP UM POR UM ==========
# for i, linha in df.iterrows():
#     if i > 0:
#         break
#     autor = limpar_nome(str(linha[0])) # Coluna A
#     titulo = limpar_nome(str(linha[1])) # Coluna B
#     link = str(linha[4]) # Coluna E

#     if not link.startswith("http"):
#         print(f"[{i}] Link inválido: {link}")
#         continue

#     print(f"\nAcessando: {link}")
#     driver.get(link)
#     time.sleep(3)  # garantir que a página carregou

#     print("Aguardando download...")
#     try:
#         arquivo_baixado = aguardar_download(download_dir, tempo_espera)
#     except Exception as e:
#         print(f" Erro ao baixar: {e}")
#         continue

#     # Novo nome + extensão
#     nova_extensao = os.path.splitext(arquivo_baixado)[-1]
#     novo_nome = f"{autor} - {titulo}{nova_extensao}"
#     caminho_destino = os.path.join(pasta_destino, novo_nome)

#     shutil.move(arquivo_baixado, caminho_destino)
#     print(f" Projeto salvo: {caminho_destino}")

#     time.sleep(3)  # esperar antes do próximo

# driver.quit()