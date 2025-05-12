import os
import shutil
from xml.dom import minidom
from openpyxl import Workbook
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import tZ13  # módulo com a função processar_dados

# Caminhos das pastas e arquivos
caminho_pasta = "C:\\Importador_CTE\\LER_XML"
#caminho_pasta_lidos = "C:\\Importador_CTE\\LIDOS_XML"
caminho_excel = "C:\\Importador_CTE\\CONSULTA_TABELAS.xlsx"
#log_path = "C:\\CHAVES_XML\\processamento_log.txt"
caminho_complementos = "C:\\Importador_CTE\\COMPLEMENTOS"
caminho_normais = "C:\\Importador_CTE\\NORMAL"

# === LIMPEZA INICIAL ===
# Remove todos os .xml nas pastas NORMAL e COMPLEMENTOS
for pasta in (caminho_normais, caminho_complementos):
    if os.path.isdir(pasta):
        for nome in os.listdir(pasta):
            if nome.lower().endswith('.xml'):
                try:
                    os.remove(os.path.join(pasta, nome))
                    print(f"Deletado {nome} em {pasta}")
                except Exception as e:
                    print(f"Erro ao deletar {nome} em {pasta}: {e}")
    else:
        print(f"Pasta não existe, criando: {pasta}")
        os.makedirs(pasta, exist_ok=True)

# Listar todos os arquivos XML na pasta de leitura
arquivos = [f for f in os.listdir(caminho_pasta) if f.endswith('.xml')]
print("Arquivos encontrados:", arquivos)

# Preparar o Excel de saída
workbook = Workbook()
sheet = workbook.active
sheet.append([
    'nCT', '', 'nCT_9d', '', 'nCT_Aspas', '', 'nCT_9d_Aspas', '',
    'CNPJ Emit', '', 'CNPJ Rem', '', 'Nome_Arquivo', '', 'infCteComp'
])

# Listas de controle
arquivos_processados = []
arquivos_com_erro = []
arquivos_complementos = []  # para infCteComp = 'C'
arquivos_normais = []       # para infCteComp = 'N'

def processar_arquivo(arquivo):
    caminho_arquivo = os.path.join(caminho_pasta, arquivo)
    try:
        print(f"Processando arquivo: {arquivo}")
        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            xml = minidom.parse(f)

        # Extração das tags necessárias
        nct_tags = xml.getElementsByTagName("nCT")
        emit_cnpj_tags = xml.getElementsByTagName("emit")[0].getElementsByTagName("CNPJ")
        rem_cnpj_tags = xml.getElementsByTagName("rem")[0].getElementsByTagName("CNPJ")
        inf_cte_comp_tags = xml.getElementsByTagName("infCteComp")

        # Valores iniciais
        nct_value = ''
        emit_cnpj_value = ''
        rem_cnpj_value = ''
        inf_cte_comp_value = 'N'  # padrão

        if nct_tags and nct_tags[0].firstChild:
            nct_value = nct_tags[0].firstChild.data.strip()
        if emit_cnpj_tags and emit_cnpj_tags[0].firstChild:
            emit_cnpj_value = emit_cnpj_tags[0].firstChild.data.strip()
        if rem_cnpj_tags and rem_cnpj_tags[0].firstChild:
            rem_cnpj_value = rem_cnpj_tags[0].firstChild.data.strip()
        if inf_cte_comp_tags:
            inf_cte_comp_value = 'C'

        # Formatação dos valores
        nct_sem_aspas = nct_value.replace("'", "")
        nct_preenchido = nct_sem_aspas.zfill(9)
        nct_aspas = f"'{nct_value}',"
        nct_preenchido_aspas = f"'{nct_preenchido}',"

        linha = [
            nct_value, '', nct_preenchido, '', nct_aspas, '',
            nct_preenchido_aspas, '', emit_cnpj_value, '',
            rem_cnpj_value, '', arquivo, '', inf_cte_comp_value
        ]
        return (arquivo, linha)
    except Exception as e:
        print(f"Erro no arquivo {arquivo}: {e}")
        return (arquivo, f"Erro ao processar o arquivo {arquivo}: {e}")

# Processamento paralelo dos XMLs
with ThreadPoolExecutor(max_workers=5) as executor:
    futures = [executor.submit(processar_arquivo, arq) for arq in arquivos]
    for future in as_completed(futures):
        arquivo, resultado = future.result()
        if isinstance(resultado, list):
            sheet.append(resultado)
            arquivos_processados.append(arquivo)
            # categoriza para mover depois
            if resultado[-1] == 'C':
                arquivos_complementos.append(arquivo)
            else:
                arquivos_normais.append(arquivo)
        else:
            sheet.append([resultado])
            arquivos_com_erro.append((arquivo, resultado))

# Salvar Excel
workbook.save(caminho_excel)
print("Processamento concluído. Dados salvos em:", caminho_excel)

# Executar formatação e INSERTs via tZ13
tZ13.processar_dados(caminho_excel)

# Gravar log de processamento
#with open(log_path, 'w') as log_file:
#    for arquivo, erro in arquivos_com_erro:
#        log_file.write(f"Erro no arquivo {arquivo}: {erro}\n")
#    for arquivo in arquivos_processados:
#        log_file.write(f"Sucesso no arquivo {arquivo}\n")

# Criar pastas de saída, se não existirem
os.makedirs(caminho_complementos, exist_ok=True)
os.makedirs(caminho_normais, exist_ok=True)

# Mover arquivos de complementos (infCteComp = 'C')
for arquivo in arquivos_complementos:
    src = os.path.join(caminho_pasta, arquivo)
    dst = os.path.join(caminho_complementos, arquivo)
    try:
        shutil.move(src, dst)
        print(f"Movido complemento: {arquivo}")
    except Exception as e:
        print(f"Erro ao mover {arquivo}: {e}")

# Mover arquivos normais (infCteComp = 'N')
for arquivo in arquivos_normais:
    src = os.path.join(caminho_pasta, arquivo)
    dst = os.path.join(caminho_normais, arquivo)
    try:
        shutil.move(src, dst)
        print(f"Movido normal: {arquivo}")
    except Exception as e:
        print(f"Erro ao mover {arquivo}: {e}")

print("Processamento finalizado.")
