import os
import pdfplumber
from docx import Document

# Caminho para a pasta onde estão os arquivos PDF
caminho_pasta_pdf = 'C:\\LER_SALVAR_PDF_WORD'

# Listar todos os arquivos na pasta que terminam com '.pdf'
arquivos_pdf = [f for f in os.listdir(caminho_pasta_pdf) if f.endswith('.pdf')]

# Verificar se há arquivos PDF na pasta
if not arquivos_pdf:
    print("Nenhum arquivo PDF encontrado na pasta.")
else:
    for arquivo_pdf in arquivos_pdf:
        caminho_pdf = os.path.join(caminho_pasta_pdf, arquivo_pdf)
        
        try:
            # Criar um novo documento Word
            documento = Document()

            # Abrir e ler o conteúdo do PDF
            with pdfplumber.open(caminho_pdf) as pdf:
                for pagina in pdf.pages:
                    texto = pagina.extract_text()
                    if texto:
                        # Adicionar o texto extraído ao documento Word
                        documento.add_paragraph(texto)

            # Salvar o documento Word com o mesmo nome do PDF
            nome_arquivo_docx = os.path.splitext(arquivo_pdf)[0] + '.docx'
            caminho_docx = os.path.join(caminho_pasta_pdf, nome_arquivo_docx)
            documento.save(caminho_docx)

            print(f"O conteúdo do PDF '{arquivo_pdf}' foi salvo em '{nome_arquivo_docx}'")

        except PermissionError as e:
            print(f"Erro de permissão ao processar '{arquivo_pdf}': {e}. Verifique as permissões do diretório e tente novamente.")
        except FileNotFoundError as e:
            print(f"Erro de arquivo não encontrado ao processar '{arquivo_pdf}': {e}. Verifique o caminho do arquivo e tente novamente.")
        except Exception as e:
            print(f"Ocorreu um erro ao processar '{arquivo_pdf}': {e}")
