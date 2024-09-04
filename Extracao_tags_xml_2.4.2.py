import os
import shutil
from xml.dom import minidom
from openpyxl import Workbook
import time

# Caminho da pasta onde estão os arquivos XML
caminho_pasta = "C:\\LER_XML"
caminho_pasta_lidos = "C:\\LIDOS_XML"

# Caminho do arquivo Excel de saída
caminho_excel = "C:\\CHAVES_XML\\mark3.xlsx"

# Listar todos os arquivos na pasta
arquivos = [f for f in os.listdir(caminho_pasta) if f.endswith('.xml')]
print("Arquivos encontrados:", arquivos)

# Criar um novo arquivo Excel
workbook = Workbook()
sheet = workbook.active

# Escrever o cabeçalho no arquivo Excel (nome_arquivo agora está no final)
sheet.append(['Numero_CTE','Serie_CTE', 'Data_Emissao', 'Chave_nCT', 'CNPJ_Emit', 'xNome_Emit', 'Valor_CTE', 'CNPJ_Dest', 'xNome_Dest', 'Nome_Arquivo'])

arquivos_processados = []
arquivos_com_erro = []

# Iterar sobre cada arquivo XML na pasta
for arquivo in arquivos:
    try:
        print(f"Processando arquivo: {arquivo}")
        caminho_arquivo = os.path.join(caminho_pasta, arquivo)
        
        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            xml = minidom.parse(f)

            # Extração de tags
            nct_tags = xml.getElementsByTagName("nCT")
            serie_tags = xml.getElementsByTagName("serie") # Acrescentado Serie CTE 04/09/2024
            infDoc_tags = xml.getElementsByTagName("infDoc")
            infNFe_tags = xml.getElementsByTagName("infNFe")
            dest_tags = xml.getElementsByTagName("dest")
            emit_tags = xml.getElementsByTagName("emit")
            dhEmi_tags = xml.getElementsByTagName("dhEmi")
            vTPrest_tags = xml.getElementsByTagName("vTPrest")
            
            # Verificar se a tag nCT existe
            if nct_tags and nct_tags[0].firstChild:
                numero_value = nct_tags[0].firstChild.data.strip()
            else:
                numero_value = 'NAO POSSUI'

            # Acrescentado Serie CTE 04/09/2024
            if serie_tags [0].firstChild:
                serie_value = serie_tags[0].firstChild.data.strip() 
            else:
                serie_value = 'NAO POSSUI'    

            
            # Extrair a tag <chave> dentro de <infNFe> que está dentro de <infDoc>
            chave_value = 'NAO POSSUI'  # Definindo valor padrão
            if infDoc_tags and infDoc_tags[0].getElementsByTagName("infNFe"):
                infNFe_tag = infDoc_tags[0].getElementsByTagName("infNFe")[0]
                chave_tags = infNFe_tag.getElementsByTagName("chave")
                if chave_tags and chave_tags[0].firstChild:
                    chave_value = chave_tags[0].firstChild.data.strip()
            
            # Extração de CNPJ e xNome de dentro da tag <dest>
            cnpj_dest_value = dest_tags[0].getElementsByTagName("CNPJ")[0].firstChild.data.strip() if dest_tags and dest_tags[0].getElementsByTagName("CNPJ")[0].firstChild else 'NAO POSSUI'
            xnome_dest_value = dest_tags[0].getElementsByTagName("xNome")[0].firstChild.data.strip() if dest_tags and dest_tags[0].getElementsByTagName("xNome")[0].firstChild else 'NAO POSSUI'
            
            # Extração de CNPJ e xNome de dentro da tag <emit>
            cnpj_emit_value = emit_tags[0].getElementsByTagName("CNPJ")[0].firstChild.data.strip() if emit_tags and emit_tags[0].getElementsByTagName("CNPJ")[0].firstChild else 'NAO POSSUI'
            xnome_emit_value = emit_tags[0].getElementsByTagName("xNome")[0].firstChild.data.strip() if emit_tags and emit_tags[0].getElementsByTagName("xNome")[0].firstChild else 'NAO POSSUI'
            
            # Extração de Data de Emissão e Valor Total da Prestação
            dhEmi_value = dhEmi_tags[0].firstChild.data.strip()[:10] if dhEmi_tags and dhEmi_tags[0].firstChild else 'NAO POSSUI'
            vTPrest_value = vTPrest_tags[0].firstChild.data.strip().replace('.', ',') if vTPrest_tags and vTPrest_tags[0].firstChild else 'NAO POSSUI'
            
            # Adicionar os valores na planilha (nome do arquivo agora é o último)
            sheet.append([numero_value, serie_value,dhEmi_value, f"'{chave_value}", cnpj_emit_value, xnome_emit_value, vTPrest_value, cnpj_dest_value, xnome_dest_value, arquivo])

        # Mover o arquivo para a pasta de lidos após o processamento
        shutil.move(caminho_arquivo, os.path.join(caminho_pasta_lidos, arquivo))
        arquivos_processados.append(arquivo)
        
        time.sleep(0.1)  # Pausa para liberar recursos do sistema

    except Exception as e:
        print(f"Erro no arquivo {arquivo}: {e}")
        sheet.append([f'Erro ao processar o arquivo: {str(e)}', 'NAO POSSUI', 'NAO POSSUI', 'NAO POSSUI', 'NAO POSSUI', 'NAO POSSUI', 'NAO POSSUI', 'NAO POSSUI', arquivo])
        arquivos_com_erro.append((arquivo, str(e)))

# Verifique os arquivos que não foram processados
arquivos_nao_processados = [f for f in os.listdir(caminho_pasta) if f.endswith('.xml')]
print("Arquivos não processados:", arquivos_nao_processados)

# Salvar o arquivo Excel
workbook.save(caminho_excel)
print("Processamento concluído. Dados salvos em:", caminho_excel)

# Registrar erros em um arquivo de log
log_path = "C:\\CHAVES_XML\\processamento_log.txt"
with open(log_path, 'w') as log_file:
    for arquivo in arquivos_com_erro:
        log_file.write(f"Erro no arquivo {arquivo[0]}: {arquivo[1]}\n")
    for arquivo in arquivos_processados:
        log_file.write(f"Sucesso no arquivo {arquivo}\n")
