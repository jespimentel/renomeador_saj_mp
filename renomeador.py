import os
import re
import datetime
import zipfile
import xml.etree.ElementTree as ET

caminho = r"D:\Du\OneDrive - Ministério Público SP\__INTERNAL_SAJ_DOCS"

def extrair_xml_docx(caminho_docx):
    with zipfile.ZipFile(caminho_docx, "r") as docx_zip:
        with docx_zip.open("word/document.xml") as xml_file:
            return xml_file.read().decode("utf-8")

def limpar_tags_xml(xml_string):
    root = ET.fromstring(xml_string)
    texto = ' '.join(root.itertext())
    return texto

def encontrar_primeiro_numero_cnj(texto):
    padrao = r"\d{7}-\d{2}\.\d{4}\.8\.26\.\d{4}"
    resultado = re.search(padrao, texto)
    if resultado:
        return resultado.group(0)

def obter_data_modificacao(caminho_docx):
    data_modificacao = os.path.getmtime(caminho_docx)
    data_modificacao = datetime.datetime.fromtimestamp(data_modificacao)
    data_formatada = data_modificacao.strftime("%Y-%m-%d_%H-%M")     
    return data_formatada

def obter_tipo_peca(xml_limpo):
    tipo_peca = xml_limpo.upper().split(' ')[0]
    mapeamento = {
        'MANIFESTAÇÃO': 'cota',
        'DECISÃO': 'arq',
        'CONTRARRAZÕES': 'cr',
        'RAZÕES': 'apel',
        'ALEGAÇÕES': 'af',
        'EXCELENTÍSSIMO': 'pet',
    }
    return mapeamento.get(tipo_peca, None)                    
                        
arquivos = os.listdir(caminho)

for arquivo in arquivos:
    if arquivo.endswith('.docx'):
        caminho_completo = os.path.join(caminho, arquivo)
        try:
            data_modificacao = obter_data_modificacao(caminho_completo)
            xml_conteudo = extrair_xml_docx(caminho_completo)
            xml_limpo = limpar_tags_xml(xml_conteudo)
            cnj = encontrar_primeiro_numero_cnj(xml_limpo)
            tipo_peca = obter_tipo_peca(xml_limpo)

            #Renomeando arquivo
            novo_nome = f"{caminho}\\{cnj} - {data_modificacao} - {tipo_peca}.docx"
            os.rename(caminho_completo, novo_nome)
            print(f"{caminho_completo} renomeado para {novo_nome}")
        except Exception as e:
            print(f"Erro ao processar {arquivo}: {e}")