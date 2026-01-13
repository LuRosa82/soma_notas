import os
import xml.etree.ElementTree as ET
import fitz  # PyMuPDF para PDFs
import re

def extrair_valor_xml(caminho):
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()

        # Tenta método 1: Iterar por tags comuns
        for tag in root.iter():
            if any(term in tag.tag.lower() for term in ['vserv', 'vliquido', 'vnf', 'valortotal', 'valorservicos', 'valordocumento']):
                if tag.text:
                    return float(tag.text.replace(',', '.'))
        
        # Tenta método 2: Usar XPath para buscar o valor em qualquer lugar
        valor_xpath = root.find('.//vServ')
        if valor_xpath is not None and valor_xpath.text:
            return float(valor_xpath.text.replace(',', '.'))

    except Exception as e:
        # Linha de depuração ativada para mostrar o erro exato:
        print(f"Erro ao processar XML {caminho}: {e}")
        return 0
    return 0

def extrair_valor_pdf(caminho):
    try:
        doc = fitz.open(caminho)
        texto = "".join([page.get_text() for page in doc])
        # Busca por padrões como "Valor Total R$ 1.234,56" ou "Total do Serviço: 100,00"
        match = re.search(r'(?:TOTAL|VALOR|R\$).*?([\d\.]+,\d{2})', texto, re.IGNORECASE | re.DOTALL)
        if match:
            return float(match.group(1).replace('.', '').replace(',', '.'))
    except: return 0
    return 0

diretorio = "." # Procura na pasta atual

total = 0

for arquivo in os.listdir(diretorio):
    caminho = os.path.join(diretorio, arquivo)
    # Linha de confirmação de leitura:
    print(f"Tentando ler: {arquivo}")
    if arquivo.endswith('.xml'):
        total += extrair_valor_xml(caminho)
    elif arquivo.endswith('.pdf'):
        total += extrair_valor_pdf(caminho)

print(f"Totalizador das Notas: R$ {total:.2f}")

