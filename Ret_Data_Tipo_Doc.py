import os
import openpyxl
from openpyxl import Workbook

def find_last_occurrence(file_path, search_term):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            last_pos = content.rfind(search_term)
            if last_pos != -1:
                extracted_str = content[last_pos + len(search_term) + 2:last_pos + len(search_term) + 21]
                return extracted_str
            else:
                return None
    except Exception as e:
        print(f"Erro na função find_last_occurrence no arquivo {file_path}: {e}")
        return None

def find_keyword_in_first_30_chars(file_path, keywords):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read(30)  # Lê as primeiras 30 posições
            for keyword in keywords:
                if keyword in content:
                    return keyword
            return None
    except Exception as e:
        print(f"Erro na função find_keyword_in_first_30_chars no arquivo {file_path}: {e}")
        return None

def search_in_files(start_path, search_term, keywords):
    results = []
    for root, dirs, files in os.walk(start_path):
        for file in files:
            if file.endswith('.txt'):
                file_path = os.path.join(root, file)
                extracted_str = find_last_occurrence(file_path, search_term)
                keyword_found = find_keyword_in_first_30_chars(file_path, keywords)
                if extracted_str or keyword_found:  # Alterado para OR
                    results.append((file, root, extracted_str, keyword_found))
    return results

def save_to_xlsx(results, xlsx_file):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Results"
    sheet.append(['File Name', 'Folder Path', 'Extracted String', 'Data Doc', 'Tipo Documento'])

    for row, result in enumerate(results, start=2):
        file_name, folder_path, extracted_str, keyword_found = result
        sheet.append([file_name, folder_path, extracted_str, '', keyword_found, ''])
        sheet[f'D{row}'] = f'=TEXT(MID(C{row},1,10),"DD/MM/YYYY")'

    workbook.save(xlsx_file)

# Caminho inicial, termo de pesquisa e lista de palavras-chave
start_path = r'C:\Temp\PDFS'
search_term = 'DATE_ATOM'
keywords = ['ADITIVO','ACORDO COMERCIAL','ACORDO DE CONFIDENCIALIDADE','ADITIVO','Apólice','CONTRATO','TERMO DE CONVÊNIO','Documento de Alteração Contratual','DECLARAÇÃO DE CAPACIDADE OPERACIONAL E ADESÃO','FORMULÁRIO DE CONTRATAÇÃO','FORMULÁRIO DE CONTRATAÇÃO','INSTRUMENTO PARTICULAR DE CESSÃO DE ATIVOS','INSTRUMENTO PARTICULAR DE CESSÃO DE DIREITOS DE USO DE AERONAVES E OUTRAS AVENÇAS','INSTRUMENTO PARTICULAR DE PRESTAÇÃO DE SERVIÇOS E OUTRAS AVENÇAS','INSTRUMENTO PARTICULAR DE PRESTAÇÃO DE SERVIÇOS DE PROFISSIONAL AUTONOMO','INTRUMENTRO PARTICULAR DE PRESTAÇÃO DE SERVIÇOS','Minuta','Memorando de Entendimentos','PROPOSTA COMERCIAL','TERMO DE PATROCINIO','TERMO DE ADESÃO','TERMO DE ADESÃO','TERMO DE ANTECIPAÇÃO','TERMO DE CONTRATAÇÃO','TERMO DE RESPONSABILIDADE FINANCEIRA','TERMO DE CONTRATAÇÃO','TERMO DE CONVENIO','TERMO DE COOPERAÇÃO','TERMO DE CONDIÇÃO DE USO','TERMO DE INCENTIVO','TERMO DE PARCERIA','TERMO DE PARCERIA','Anexo','ATESTADO DE PRESTAÇÃO DE SERVIÇO','Carta','INSTRUMENTO DE DISTRATO','TERMO DE DISTRATO E QUITAÇÃO','INSTRUMENTO PARTICULAR DE CESSÃO DE USO TEMPORÁRIO DE ÁREA','INSTRUMENTO PARTICULAR DE DOAÇÃO','INSTRUMENTO PARTICULAR DE DISTRATO E QUITAÇÃO','INSTRUMENTO PARTICULAR DE QUITAÇÃO','Licença','Notificação Extrajudicial','ORDEM DE SERVIÇO','Order Form','ORDEM DE SERVIÇO','ORDEM SOLICITAÇÃO DE SERVIÇOS','PROCURAÇÃO','PRÉ-CONTRATO DE FRANQUIA','PEDIDO','Proposta Técnica','RECIBO','TERMO DE CESSÃO DE DIREITOS CREDITÓRIOS','TERMO DE CONFISSÃO DE DÍVIDA','TERMO DE DISTRATO','TERMO DE DISTRATO E QUITAÇÃO','termo de entrega provisória de obra','TERMO DE TRANSFERÊNCIA DE TITULARIDADE DE ACESSO DO SERVIÇO MÓVEL PESSOAL','Termo de Recebimento de Chaves','TERMO PARTICULAR DE DOAÇÃO','TERMO DE QUITAÇÃO','TERMO DE RESPONSABILIDADE','Termo de Solicitação de Serviços','TERMO DE TRANSFERÊNCIA DE TITULARIDADE','INSTRUMENTO PARTICULAR DE CESSÃO DA TITULARIDADE DO CÓDIGO DE ACESSO DO SEVIÇO MÓVEL PESSOAL (SMP)']

# Realiza a busca e salva os resultados
results = search_in_files(start_path, search_term, keywords)
save_to_xlsx(results, r'C:\Temp\PDFS\DocData.xlsx')

if results:
    print("Processo concluído. Os resultados foram salvos em 'C:\Temp\PDFS\DocData.xlsx'.")
else:
    print("Nenhum resultado encontrado.")
