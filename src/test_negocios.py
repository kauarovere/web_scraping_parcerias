import requests
from bs4 import BeautifulSoup
import urllib.parse
import unicodedata
import re

def limpar_rotulo(texto):
    if not texto: return ""
    texto = str(texto).lower().replace(":", "")
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return " ".join(texto.split())

def extrair_informacoes_hibrido(html_content):
    soup  = BeautifulSoup(html_content, 'html.parser')
    linhas = list(soup.stripped_strings)

    modalidade  = "Não identificada"
    data_inicio = "Não identificada"
    data_fim    = "Não identificada"

    for i, linha in enumerate(linhas):
        linha_original = linha.strip()
        linha_limpa    = limpar_rotulo(linha_original)

        if "modalidade" in linha_limpa and modalidade == "Não identificada":
            if linha_limpa == "modalidade":
                if i + 1 < len(linhas): modalidade = linhas[i+1].strip()
            else:
                modalidade = linha_original.lower().replace("modalidade", "").replace(":", "").strip().title()

        elif "data de inicio" in linha_limpa and data_inicio == "Não identificada":
            if linha_limpa == "data de inicio":
                if i + 1 < len(linhas): data_inicio = linhas[i+1].strip()
            else:
                data_inicio = linha_original.split(":")[-1].strip() if ":" in linha_original else linha_original.replace("Data de Início", "").strip()

        elif "data de fim" in linha_limpa and data_fim == "Não identificada":
            if linha_limpa == "data de fim":
                if i + 1 < len(linhas): data_fim = linhas[i+1].strip()
            else:
                data_fim = linha_original.split(":")[-1].strip() if ":" in linha_original else linha_original.replace("Data de Fim", "").strip()

    return modalidade, data_inicio, data_fim

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/148.0.0.0 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded',
}

session = requests.Session()
url_base = 'https://diariooficial.prefeitura.sp.gov.br/md_epubli_controlador.php?acao=negocios_pesquisar'

processo = "6016.2020/0101272-2"
payload = {
    'hdnObjeto': processo,
    'hdnInicio': '0',
    'hdnModoPesquisa': 'DATA'
}

res_post = session.post(url_base, headers=headers, data=payload)
soup_post = BeautifulSoup(res_post.text, 'html.parser')
links = soup_post.find_all('a', href=True)
doc_links = [a for a in links if 'md_epubli_visualizar' in a['href']]

if len(doc_links) > 1: # Index 0 usually is the global panel desc, docs start at 1
    target_link = doc_links[1]['href']
    if not target_link.startswith('http'):
        target_link = "https://diariooficial.prefeitura.sp.gov.br/" + target_link
    
    print("Testando link:", target_link)
    resp_doc = session.get(target_link, headers=headers)
    if 'iso-8859-1' in resp_doc.text.lower():
        resp_doc.encoding = 'iso-8859-1'
    
    mod, d_ini, d_fim = extrair_informacoes_hibrido(resp_doc.text)
    print(f"Modalidade: {mod}")
    print(f"Data Início: {d_ini}")
    print(f"Data Fim: {d_fim}")
