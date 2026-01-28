import os
import xml.etree.ElementTree as ET
import fitz  # PyMuPDF
import re
import pandas as pd # type: ignore

def extrair_valor_xml(caminho):
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()
        # Remove namespaces para facilitar a busca
        for el in root.iter():
            if '}' in el.tag:
                el.tag = el.tag.split('}', 1)[1]
        
        # Lista de prioridade de tags de valor
        tags_valor = ['vNF', 'vServ', 'vLiquido', 'vLiq', 'vTotal']
        for tag_nome in tags_valor:
            alvo = root.find(f".//{tag_nome}")
            if alvo is not None and alvo.text:
                return float(alvo.text)
    except Exception as e:
        print(f"Erro XML {caminho}: {e}")
    return 0.0

def extrair_valor_pdf(caminho):
    try:
        doc = fitz.open(caminho)
        # Pega apenas a primeira página para evitar valores de cobranças futuras
        texto = doc[0].get_text()
        
        # Regex melhorado: busca valor após termos-chave, ignorando labels comuns
        padrao = r'(?:VALOR TOTAL|TOTAL DA NOTA|VALOR LÍQUIDO).*?(\d{1,3}(?:\.\d{3})*,\d{2})'
        match = re.search(padrao, texto, re.IGNORECASE | re.DOTALL)
        
        if match:
            valor_str = match.group(1).replace('.', '').replace(',', '.')
            return float(valor_str)
    except Exception as e:
        print(f"Erro PDF {caminho}: {e}")
    return 0.0

def processar_pasta(diretorio="."):
    dados = []
    
    for arquivo in os.listdir(diretorio):
        caminho = os.path.join(diretorio, arquivo)
        valor = 0
        
        if arquivo.endswith('.xml'):
            valor = extrair_valor_xml(caminho)
        elif arquivo.endswith('.pdf'):
            valor = extrair_valor_pdf(caminho)
        else:
            continue
            
        dados.append({'Arquivo': arquivo, 'Valor': valor})
        print(f"Processado: {arquivo} -> R$ {valor:.2f}")

    # Criar DataFrame e salvar relatório
    df = pd.DataFrame(dados)
    df.to_csv("relatorio_notas.csv", index=False, sep=';', encoding='utf-8-sig')
    
    print("\n" + "="*30)
    print(f"TOTAL GERAL: R$ {df['Valor'].sum():.2f}")
    print(f"Relatório gerado: relatorio_notas.csv")
    print("="*30)

if __name__ == "__main__":
    processar_pasta()
