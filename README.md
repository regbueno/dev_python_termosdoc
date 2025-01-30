# dev_python_termosdoc
Pesquisa de termos em documentos PDF
Funcionalidades:
Upload de Documentos:
● Permitir que o usuário faça o upload de um ou mais documentos nos formatos comuns (.doc, .docx, .pdf, .txt).
● Verifi cação de integridade do documento após o upload.
● Capacidade de visualizar uma lista de todos os documentos carregados.
● Verifi car se o documento é uma imagem. Caso positivo, transformar em texto, utilizando OCR.
● Converta todos os documentos em TXT
Defi nição de Palavras-Chave:
● Permitir que o usuário digite termos de pesquisa.
● Os termos digitados devem ser pesquisados nos arquivos TXT
● Os termos digitados devem ser reconhecidos independente se digitados em maiúsculo ou minúsculo.
Análise de Conteúdo:
● Realizar uma varredura no conteúdo dos documentos em busca dos termos digitados.
● Programa deve analisar a quantidade de vezes que cada termo aparece em cada documento.
Relatório de Análise:
● Gerar um relatório/planilha detalhado contendo:
■ Nome do Documento;
■ Endereço da Pasta
■ Termo 1;
■ Termo 2;
■ Termo n.
● Na planilha gerada, abaixo de cada termo, deve ser apresentado a quantidade de vezes que o termo digitado foi encontrado.
Ex:
Nome do Arquivo
Pasta
QTD-anoimiza
QTD-dados pessoais
QTD_lgpd
QTD-privacidade
Flores.PDF
C:\Temp
5
0
1
2
Carros.DOC
C:\Temp
4
0
0
0
Casas.PDF
C:\Temp
2
0
1
3
Filtros e Pesquisa:
● Fornecer opções de fi ltro para visualizar documentos específi cos ou palavras-chave específi cas.
● Capacidade de pesquisa por nome de documento, conteúdo ou palavra-chave.
Interface Amigável:
● Design intuitivo e fácil de navegar.
● Feedback visual durante a análise de documentos.
Segurança:
● Criptografi a para o armazenamento de documentos.
● Proteção contra uploads de arquivos maliciosos.
● Opção de autenticação por SSO.
Abaixo, segue exemplo em linguagem natural em Português
1. Preparação e Confi guração:
● Importação de Bibliotecas: Importa as bibliotecas necessárias para manipulação de arquivos ( os ), processamento de imagens e OCR ( pytesseract, pdf2image ), manipulação de dados ( pandas ), normalização de texto ( unidecode ), processamento de linguagem natural ( spacy, nltk ).
● Confi guração do Tesseract OCR: Defi ne o caminho para o executável do Tesseract, essencial para a conversão de imagens em texto.
● Carregamento do Modelo Spacy: Carrega o modelo de linguagem en_core_web_sm do Spacy, usado para processamento de texto.
● Download de Recursos do NLTK: Baixa o recurso wordnet do NLTK, usado para encontrar sinônimos.
2. Funções Auxiliares:
● fi nd_synonyms(word):
○ Objetivo: Encontrar sinônimos de uma palavra.
○ Funcionamento: Usa o wordnet do NLTK para buscar sinônimos de uma palavra. Ele itera pelos synsets (conjuntos de sinônimos) e lemmas (formas lexicais) associados à palavra, adicionando-os a um conjunto de sinônimos. Os sinônimos são então retornados como uma lista.
● pdf_to_text(pdf_path):
○ Objetivo: Converter um arquivo PDF em texto.
○ Funcionamento: Usa a biblioteca pdf2image para converter cada página do PDF em uma imagem. Em seguida, usa o pytesseract para realizar OCR (Reconhecimento Óptico de Caracteres) em
cada imagem e extrair o texto. O texto de todas as páginas é concatenado e retornado.
● normalize_text(text):
○ Objetivo: Normalizar o texto.
○ Funcionamento: Remove acentos e converte o texto para minúsculas usando a biblioteca unidecode.
● search_terms_in_text(text, terms):
○ Objetivo: Buscar termos e seus sinônimos no texto e contar suas ocorrências.
○ Funcionamento: Usa o Spacy para processar o texto. Para cada termo de busca fornecido, ele encontra seus sinônimos usando a função fi nd_synonyms. Em seguida, conta quantas vezes o termo e seus sinônimos aparecem no texto e retorna um dicionário com o termo como chave e a contagem total como valor.
● convert_pdfs_in_folder(root_folder):
○ Objetivo: Converter todos os PDFs de uma pasta e suas subpastas para arquivos TXT.
○ Funcionamento: Percorre recursivamente a pasta raiz e suas subpastas. Para cada arquivo PDF encontrado, ele converte o PDF para texto usando pdf_to_text, normaliza o texto com normalize_text e salva o texto normalizado em um arquivo TXT com o mesmo nome do PDF, mas com a extensão .txt.
3. Execução Principal:
● Defi nição da Pasta Raiz: Defi ne a pasta C:\\TEMP\\PDFS como a pasta raiz onde os PDFs estão localizados.
● Conversão de PDFs para TXT: Chama a função convert_pdfs_in_folder para converter todos os PDFs na pasta raiz e subpastas para arquivos TXT.
● Entrada do Usuário: Solicita ao usuário que insira os termos de busca, separados por vírgula.
● Processamento dos Termos de Busca: Normaliza os termos de busca (remove acentos e converte para minúsculas).
● Inicialização dos Resultados: Cria um dicionário para armazenar os resultados da busca, com cada termo de busca como chave e uma lista vazia como valor.
● Busca nos Arquivos TXT: Percorre novamente a pasta raiz e suas subpastas. Para cada arquivo TXT, ele lê o conteúdo, busca os termos e seus sinônimos usando search_terms_in_text e atualiza o dicionário de resultados com o nome do arquivo, a pasta e a contagem de cada termo encontrado.
● Preparação dos Dados para a Planilha: Formata os resultados em um dicionário que pode ser facilmente convertido em um DataFrame do Pandas. Ele cria colunas para o nome do arquivo, pasta e a contagem de cada termo.
● Criação da Planilha: Cria um DataFrame do Pandas a partir do dicionário de dados e salva o DataFrame em uma planilha Excel chamada ANALISE_TERMOS.xlsx na pasta raiz.
● Mensagem de Conclusão: Imprime uma mensagem informando que a análise foi concluída e o nome do arquivo onde os resultados foram salvos.
Outra funcionalidade importante é gerar relação dos tipos de documentos e suas pastas.
1. Importação de Bibliotecas:
● os: Fornece funções para interagir com o sistema operacional, como navegar por pastas e manipular arquivos.
● openpyxl: Permite criar e manipular arquivos Excel (.xlsx).
2. Funções:
● find_last_occurrence(file_path, search_term):
○ Objetivo: Encontrar a última ocorrência de um termo de pesquisa em um arquivo e extrair uma string próxima a ele.
○ Funcionamento:
■ Abre o arquivo especifi cado em file_path em modo leitura ('r') e com codifi cação utf-8.
■ Lê todo o conteúdo do arquivo para a variável content.
■ Usa o método rfind() para encontrar a última posição (índice) do search_term dentro de content.
■ Se o termo for encontrado (ou seja, se last_pos for diferente de -1):
■ Extrai uma substring de 20 caracteres que começa 2 posições após o fi nal do search_term. Essa substring é armazenada em extracted_str.
■ Retorna a extracted_str.
■ Se o termo não for encontrado, retorna None.
● find_keyword_in_first_30_chars(file_path, keywords):
○ Objetivo: Verifi car se alguma palavra-chave de uma lista está presente nos primeiros 30 caracteres de um arquivo.
○ Funcionamento:
■ Abre o arquivo especifi cado em file_path em modo leitura ('r') e com codifi cação utf-8.
■ Lê os primeiros 30 caracteres do arquivo e armazena em content.
■ Itera sobre a lista de keywords.
■ Para cada keyword, verifi ca se ela está contida em content.
■ Se uma palavra-chave for encontrada, retorna a palavra-chave.
■ Se nenhuma palavra-chave for encontrada, retorna None.
● search_in_files(start_path, search_term, keywords):
○ Objetivo: Percorrer uma pasta e suas subpastas, buscar arquivos de texto que contenham um termo específi co, extrair informações relevantes e identifi car uma palavra-chave nos primeiros 30 caracteres.
○ Funcionamento:
■ Inicializa uma lista vazia results para armazenar os resultados.
■ Usa os.walk(start_path) para percorrer recursivamente a pasta start_path e suas subpastas.
■ Para cada arquivo encontrado:
■ Verifi ca se o arquivo termina com .txt.
■ Se for um arquivo de texto, constrói o caminho completo do arquivo (file_path).
■ Chama find_last_occurrence para extrair a string próxima à última ocorrência do search_term.
■ Chama find_keyword_in_first_30_chars para buscar uma palavra-chave nos primeiros 30 caracteres do arquivo.
■ Se ambas as funções retornarem valores válidos (ou seja, extracted_str e keyword_found não forem None):
■ Adiciona uma tupla contendo o nome do arquivo, a pasta, a string extraída e a palavra-chave encontrada à lista results.
■ Retorna a lista results.
● save_to_xlsx(results, xlsx_file):
○ Objetivo: Salvar os resultados da busca em uma planilha Excel.
○ Funcionamento:
■ Cria uma nova pasta de trabalho do Excel usando Workbook().
■ Seleciona a planilha ativa (active).
■ Defi ne o título da planilha como "Results".
■ Adiciona um cabeçalho à planilha com os seguintes títulos: 'File Name', 'Folder Path', 'Extracted String', 'Data Doc', 'Tipo Documento'.
■ Itera sobre os resultados em results, começando da linha 2 (para pular o cabeçalho).
■ Para cada resultado, extrai o nome do arquivo, a pasta, a string extraída e a palavra-chave encontrada.
■ Adiciona uma nova linha à planilha com essas informações.
■ Aplica uma fórmula Excel na coluna 'Data Doc' (coluna D) para formatar a string extraída como uma data no formato DD/MM/YYYY. A fórmula assume que os 10 primeiros caracteres da string extraída representam uma data no formato YYYY-MM-DD.
■ Salva a pasta de trabalho no arquivo especifi cado por xlsx_file.
3. Execução do Código:
● Defi nição de Variáveis:
○ start_path: Defi ne o caminho da pasta onde a busca será iniciada ( r'C:\Temp\PDFS' ).
○ search_term: Defi ne o termo que será buscado nos arquivos ( 'DATE_ATOM' ).
○ keywords: Defi ne uma lista de palavras-chave que serão buscadas nos primeiros 30 caracteres dos arquivos.
● Chamada das Funções:
○ results = search_in_files(start_path, search_term, keywords): Chama a função search_in_files para realizar a busca nos arquivos.
○ save_to_xlsx(results, r'C:\Temp\PDFS\DocData.xlsx'): Chama a função save_to_xlsx para salvar os resultados em uma planilha Excel chamada DocData.xlsx na pasta C:\Temp\PDFS.
● Mensagem de Conclusão:
○ print("Processo concluído. Os resultados foram salvos em 'C:\Temp\PDFS\DocData.xlsx'."): Imprime uma mensagem informando que o processo foi concluído e o nome do arquivo onde os resultados foram salvos.
Em resumo, o código realiza as seguintes tarefas:
1. Busca por arquivos de texto em uma pasta e suas subpastas.
2. Encontra a última ocorrência de um termo específi co ('DATE_ATOM') nesses arquivos.
3. Extrai uma string de 20 caracteres próxima a essa ocorrência.
4. Verifi ca se alguma palavra-chave de uma lista predefi nida está presente nos primeiros 30 caracteres do arquivo.
5. Armazena os resultados (nome do arquivo, pasta, string extraída e palavra-chave) em uma lista.
6. Cria uma planilha Excel e salva os resultados, formatando a string extraída como uma data na coluna 'Data Doc'.
