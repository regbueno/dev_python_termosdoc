
# Preparação das Bibliotecas
import os
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import unidecode
import spacy
import nltk
from nltk.corpus import wordnet

# Configurar o caminho do Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Carregar modelo Spacy
nlp = spacy.load("en_core_web_sm")

# Baixar recursos do NLTK
nltk.download('wordnet')

# Função para encontrar sinônimos usando Spacy e NLTK
def find_synonyms(word):
    synonyms = set()
    for syn in wordnet.synsets(word):
        for lemma in syn.lemmas():
            synonyms.add(lemma.name().replace('_', ' '))
    return list(synonyms)

# Função para converter PDF para texto
def pdf_to_text(pdf_path):
    images = convert_from_path(pdf_path)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img)
    return text

# Função para normalizar texto
def normalize_text(text):
    return unidecode.unidecode(text).lower()

# Função para buscar termos e seus sinônimos no texto
def search_terms_in_text(text, terms):
    doc = nlp(text)
    found_terms = {}
    for term in terms:
        synonyms = find_synonyms(term)
        total_count = sum([text.count(synonym) for synonym in synonyms])
        found_terms[term] = total_count
    return found_terms

# Função para converter PDFs para TXT em uma pasta e subpastas
def convert_pdfs_in_folder(root_folder):
    for subdir, _, files in os.walk(root_folder):
        for file in files:
            if file.endswith(".pdf"):
                pdf_path = os.path.join(subdir, file)
                text = pdf_to_text(pdf_path)
                normalized_text = normalize_text(text)
                with open(pdf_path.replace(".pdf", ".txt"), "w") as text_file:
                    text_file.write(normalized_text)

# Caminho da pasta raiz
root_folder_path = "C:\\TEMP\\PDFS"

# Converter todos os PDFs para TXT nas subpastas
convert_pdfs_in_folder(root_folder_path)

# Pedir ao usuário múltiplos termos para buscar
search_terms_input = input("Quais termos devo procurar? (separe por vírgula) ")
search_terms = [unidecode.unidecode(term.strip()).lower() for term in search_terms_input.split(',')]

# Inicializar o dicionário para armazenar os resultados
results = {term: [] for term in search_terms}

# Encontrar os termos e sinônimos nos arquivos TXT
for subdir, _, files in os.walk(root_folder_path):
    for file in files:
        if file.endswith(".txt"):
            file_path = os.path.join(subdir, file)
            with open(file_path, "r") as text_file:
                text = text_file.read()
                found_terms = search_terms_in_text(text, search_terms)
                for term, count in found_terms.items():
                    if count > 0:
                        results[term].append({"Nome do Arquivo": file, "Pasta": subdir, "QTD": count})

# Preparar os dados para a planilha
data = {"Nome do Arquivo": [], "Pasta": []}
for term, counts in results.items():
    data["QTD_" + term] = [count["QTD"] for count in counts]
    if not data["Nome do Arquivo"]:
        data["Nome do Arquivo"] = [count["Nome do Arquivo"] for count in counts]
        data["Pasta"] = [count["Pasta"] for count in counts]

# Criar uma planilha com os resultados
df = pd.DataFrame(data)
xls_filename = "ANALISE_TERMOS.xlsx"
df.to_excel(os.path.join(root_folder_path, xls_filename), index=False)

print(f"Análise concluída. Resultados salvos em {xls_filename}.")
