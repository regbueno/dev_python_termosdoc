import os
import pandas as pd
import docx

def ler_txt(file_path):
    with open(file_path, 'r') as file:
        return file.read()

def ler_docx(file_path):
    doc = docx.Document(file_path)
    full_text = [para.text for para in doc.paragraphs]
    return '\n'.join(full_text)

def buscar_termos(texto, termos):
    return [termo for termo in termos if termo in texto]

def processar_arquivos(diretorio, termos):
    resultados = []
    for root, dirs, files in os.walk(diretorio):
        for file in files:
            file_path = os.path.join(root, file)
            if file.endswith('.txt'):
                texto = ler_txt(file_path)
            elif file.endswith('.docx') or file.endswith('.doc'):
                texto = ler_docx(file_path)
            else:
                continue

            termos_encontrados = buscar_termos(texto, termos)
            if termos_encontrados:
                resultados.append({
                    "Nome do Arquivo": file,
                    "Caminho": file_path,
                    "Termos Encontrados": ", ".join(termos_encontrados)
                })

    return pd.DataFrame(resultados)

# Caminho para a pasta onde os arquivos estão localizados e onde a planilha será salva
diretorio = 'C:\\Temp\\Docs'
termos = ["LGPD", "Contrato", "Sensível", "Proteção de Dados"]
df = processar_arquivos(diretorio, termos)

# Salvando a planilha no mesmo diretório
caminho_planilha = os.path.join(diretorio, 'resultados.xlsx')
df.to_excel(caminho_planilha, index=False)
