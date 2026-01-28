import os
import xml.etree.ElementTree as ET
import fitz
import re
import pandas as pd
from pathlib import Path
import unicodedata
import shutil # Importa o módulo para mover arquivos

# --- FUNÇÕES AUXILIARES ---

def limpar_nome_arquivo(nome):
    """Remove caracteres inválidos para nomes de arquivos no Windows."""
    nome = ''.join(c for c in nome if unicodedata.category(c) != 'C')
    for char in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
        nome = nome.replace(char, '_')
    return nome.strip()

# --- FUNÇÕES DE EXTRAÇÃO (XML e PDF) ---

def extrair_dados_xml(caminho):
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()
        for el in root.iter():
            if '}' in el.tag: el.tag = el.tag.split('}', 1)
        emitente = root.findtext(".//PrestadorServico/RazaoSocial") or root.findtext(".//Prestador/xNome") or root.findtext(".//emit/xNome") or root.findtext(".//RazaoSocial")
        cnpj = root.findtext(".//PrestadorServico//CNPJ") or root.findtext(".//Prestador//CNPJ") or root.findtext(".//emit/CNPJ") or root.findtext(".//Cnpj")
        data = root.findtext(".//DataEmissao") or root.findtext(".//dhEmi") or root.findtext(".//dtEmi")
        tags_valor = ["vLiq", "vServ", "ValorServicos", "vNF", "vLiquido", "ValorLiquido", "vTotalRet"]
        valor_final = 0.0
        for tag_nome in tags_valor:
            valor_texto = root.findtext(f".//{tag_nome}")
            if valor_texto:
                try:
                    valor_final = float(valor_texto.replace(',', '.'))
                    if valor_final > 0.0: break
                except ValueError:
                    continue
        return {
            'Arquivo': os.path.basename(caminho),
            'Emitente': emitente if emitente else "Não encontrado",
            'CNPJ': cnpj if cnpj else "Não encontrado",
            'Data': data[:10] if data else "---",
            'Valor': valor_final
        }
    except Exception as e:
        # print(f"ERRO ao processar o arquivo {os.path.basename(caminho)}: {e}") 
        return None

def extrair_dados_pdf(caminho):
    try:
        doc = fitz.open(caminho)
        texto = " ".join([page.get_text() for page in doc])
        cnpj_match = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', texto)
        valor_match = re.search(r'(?:TOTAL|VALOR|R\$).*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto, re.IGNORECASE | re.DOTALL)
        data_match = re.search(r'(\d{2}/\d{2}/\d{4})', texto)
        return {
            'Arquivo': os.path.basename(caminho),
            'Emitente': "Verificar PDF",
            'CNPJ': cnpj_match.group(0) if cnpj_match else "Não encontrado",
            'Data': data_match.group(0) if data_match else "---",
            'Valor': float(valor_match.group(1).replace('.', '').replace(',', '.')) if valor_match else 0.0
        }
    except Exception as e: 
        # print(f"ERRO ao processar o arquivo PDF {os.path.basename(caminho)}: {e}") 
        return None

# --- LÓGICA PRINCIPAL (ONDE TUDO ACONTECE) ---

def processar_pasta(diretorio="."):
    lista_dados = []
    print(f"Iniciando processamento e organização na pasta: {Path(diretorio).resolve()}")
    
    for arquivo in os.listdir(diretorio):
        caminho_origem = os.path.join(diretorio, arquivo)
        
        if arquivo in ["leitor_notas.py", "relatorio_detalhado.csv"] or os.path.isdir(caminho_origem):
            continue

        resultado = None
        extensao = Path(arquivo).suffix.lower()
        
        if extensao == '.xml':
            resultado = extrair_dados_xml(caminho_origem)
        elif extensao == '.pdf':
            resultado = extrair_dados_pdf(caminho_origem)
        
        if resultado and resultado['Valor'] > 0 and resultado['Data'] != '---':
            lista_dados.append(resultado)
            print(f"Lido: {resultado['Arquivo']} | R$ {resultado['Valor']:.2f}")

            # Lógica de RENOMEAÇÃO
            data_obj = pd.to_datetime(resultado['Data']).date()
            mes_nome = data_obj.strftime('%m - %B') # Ex: "12 - December"
            ano = data_obj.strftime('%Y')
            
            # Formato do nome do arquivo
            data_str = data_obj.strftime('%Y%m%d')
            nome_emitente_limpo = limpar_nome_arquivo(resultado['Emitente'])
            valor_str = f"{resultado['Valor']:.2f}".replace('.', ',')
            novo_nome = f"{data_str}_{nome_emitente_limpo}_R${valor_str}{extensao}"

            # Lógica de ORGANIZAÇÃO EM PASTAS
            caminho_destino_pasta = Path(diretorio) / ano / mes_nome
            caminho_destino_pasta.mkdir(parents=True, exist_ok=True) # Cria a pasta se não existir

            caminho_destino_arquivo = caminho_destino_pasta / novo_nome

            try:
                # Usa shutil.move para renomear E mover de uma vez
                shutil.move(caminho_origem, caminho_destino_arquivo)
                print(f"  Organizado para: {caminho_destino_arquivo}")
            except shutil.Error as e:
                print(f"  Aviso: Arquivo já existe no destino, ignorando movimento. {e}")
            except Exception as e:
                print(f"  Erro ao mover {arquivo}: {e}")


    if lista_dados:
        df = pd.DataFrame(lista_dados)
        df.to_csv("relatorio_detalhado.csv", index=False, sep=';', encoding='utf-8-sig')
        total_geral = df['Valor'].sum()
        print("\n" + "="*30)
        print(f"TOTAL GERAL: R$ {total_geral:.2f}")
        print(f"Relatório 'relatorio_detalhado.csv' gerado com sucesso!")
        print("="*30)
    else:
        print("Nenhum arquivo XML ou PDF original encontrado para processar (eles podem já estar organizados em subpastas).")

if __name__ == "__main__":
    processar_pasta()
