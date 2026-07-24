import requests
from bs4 import BeautifulSoup
import re
import urllib.parse
import unicodedata
import pandas as pd
import time 
import tkinter as tk
from tkinter import filedialog

# --- 1. SELEÇÃO DO ARQUIVO COM O MOUSE ---
print("Iniciando o robô...")

root = tk.Tk()
root.withdraw() 
root.attributes('-topmost', True) 

print("Por favor, selecione o arquivo CSV ou Excel na janela que acabou de abrir...")

caminho_arquivo = filedialog.askopenfilename(
    title="Selecione o arquivo com a coluna codProcesso_y",
    filetypes=[("Arquivos CSV", "*.csv"), ("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
)

if not caminho_arquivo:
    print("❌ Nenhum arquivo selecionado. Encerrando o robô.")
    exit()

print(f"Arquivo selecionado com sucesso!")

try:
    if caminho_arquivo.lower().endswith(('.xlsx', '.xls')):
        df_entrada = pd.read_excel(caminho_arquivo)
    else:
        df_entrada = pd.read_csv(caminho_arquivo, sep=None, engine='python', encoding='latin-1')
    
    # O PULO DO GATO: Removido o .head(100). Agora ele pega a coluna inteira!
    PROCESSOS = df_entrada['codProcesso_y'].dropna().astype(str).tolist()
    print(f"✅ {len(PROCESSOS)} processos carregados com sucesso! Preparando motores...\n")
    
except Exception as e:
    print(f"❌ Erro ao ler o arquivo selecionado: {e}")
    print("Dica: Se a coluna tiver outro nome na planilha atual, o script não vai encontrá-la.")
    exit()

# --- CONFIGURAÇÕES DO ROBÔ ---
TIPOS_DOCUMENTO = ["extrato", "retificação", "alteração", "aditamento", "prorrogação", "termo de fomento", "termo de colaboração"]

url_busca = 'https://diariooficial.prefeitura.sp.gov.br/md_epubli_controlador.php?acao=materias_pesquisar'

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/148.0.0.0 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded',
}

def formatar_processo(num):
    num = re.sub(r'\D', '', num) 
    if len(num) == 16:
        return f"{num[:4]}.{num[4:8]}/{num[8:15]}-{num[15:]}"
    return num

def limpar_rotulo(texto):
    if not texto: return ""
    texto = str(texto).lower().replace(":", "")
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return " ".join(texto.split()) 

def extrair_informacoes_hibrido(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    linhas = list(soup.stripped_strings)
    
    modalidade = "Não identificada"
    data_inicio = "Não identificada"
    data_fim = "Não identificada" 
    
    for i, linha in enumerate(linhas):
        linha_original = linha.strip()
        linha_limpa = limpar_rotulo(linha_original)
        
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

# --- EXECUÇÃO EM LOTE ---
print("-" * 60)
print(f"INICIANDO EXTRAÇÃO DE {len(PROCESSOS)} PROCESSOS (BASE COMPLETA)")
print("-" * 60)

resultados_excel = []
processos_com_sucesso = 0
processos_sem_sucesso = 0

contador = 1

for processo_bruto in PROCESSOS:
    processo_fmt = formatar_processo(processo_bruto)
    print(f"[{contador}/{len(PROCESSOS)}] Buscando processo: {processo_fmt}...", end=" ")
    
    processo_codificado = urllib.parse.quote(processo_fmt)
    payload_bruto = f"hdnTermoPesquisa={processo_codificado}&hdnTipoPesquisa=Q&hdnVersaoDiario=&hdnOndePesquisa=&hdnTipoDataPesquisa=I&hdnDataInicioPesquisa=&hdnDataFimPesquisa=&hdnTipoDocumentoPesquisa=&hdnVeiculoPublicacao=&hdnDataPublicacao=&hdnOrgaoFiltro=&hdnUnidadeResponsavelFiltro=&hdnTipoDocumentoFiltro=&hdnInicio=0&hdnVisualizacao=L&hdnModoPesquisa=RAPIDA"

    response = requests.post(url_busca, headers=headers, data=payload_bruto)
    
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        
        documento_sucesso_link = None
        mod_final = "Não identificada"
        dt_ini_final = "Não identificada"
        dt_fim_final = "Não identificada"
        
        for tag_a in soup.find_all('a', href=True):
            href = tag_a['href']
            
            if 'md_epubli' in href or 'materia' in href or 'acao=exibir' in href:
                if "pesquisar" not in href and "login" not in href:
                    if not href.startswith('http'):
                        href = "https://diariooficial.prefeitura.sp.gov.br/" + href
                    
                    container = tag_a.find_parent(['div', 'p', 'tr', 'li'])
                    texto_contexto = container.text.lower() if container else tag_a.parent.text.lower()
                    
                    if any(tipo in texto_contexto for tipo in TIPOS_DOCUMENTO):
                        
                        resp_doc = requests.get(href, headers=headers)
                        if 'iso-8859-1' in resp_doc.text.lower():
                            resp_doc.encoding = 'iso-8859-1'
                            
                        mod, d_ini, d_fim = extrair_informacoes_hibrido(resp_doc.text)
                        
                        if d_ini != "Não identificada" or d_fim != "Não identificada":
                            documento_sucesso_link = href
                            mod_final = mod
                            dt_ini_final = d_ini
                            dt_fim_final = d_fim
                            break 
        
        if documento_sucesso_link:
            print("✅ OK")
            processos_com_sucesso += 1
            resultados_excel.append({
                "Nº do Processo": processo_fmt,
                "Modalidade": mod_final,
                "Data Início": dt_ini_final,
                "Data Fim": dt_fim_final,
                "Status": "Mapeado",
                "Link do Documento": documento_sucesso_link
            })
        else:
            print("⚠️ Sem dados")
            processos_sem_sucesso += 1
            resultados_excel.append({
                "Nº do Processo": processo_fmt,
                "Modalidade": "-",
                "Data Início": "-",
                "Data Fim": "-",
                "Status": "Não mapeado (Rótulos ausentes)",
                "Link do Documento": "-"
            })
            
    else:
        print("❌ Erro")
        processos_sem_sucesso += 1
        resultados_excel.append({
            "Nº do Processo": processo_fmt,
            "Modalidade": "Erro de Conexão",
            "Data Início": "-",
            "Data Fim": "-",
            "Status": "Falha na busca",
            "Link do Documento": "-"
        })
        
    contador += 1
    time.sleep(1) # Mantemos esse respiro de 1 segundo para o firewall não nos derrubar

# --- GERAÇÃO DO ARQUIVO ---
print("\n" + "=" * 60)
print("Gerando planilha Excel final...")

df_resultados = pd.DataFrame(resultados_excel)
nome_arquivo = "Auditoria_Completa_Parcerias.xlsx"
df_resultados.to_excel(nome_arquivo, index=False)

# --- INSIGHTS FINAIS ---
print("=" * 60)
print("📊 INSIGHTS DA EXTRAÇÃO COMPLETA")
print("=" * 60)
print(f"-> {processos_com_sucesso} processos foram mapeados com sucesso.")
print(f"-> {processos_sem_sucesso} processos não foram mapeados com sucesso.")
print(f"\n✅ Arquivo final salvo como: '{nome_arquivo}'")