import requests
from bs4 import BeautifulSoup
import re
import unicodedata
import pandas as pd
import time
import json
import os
import datetime
import tkinter as tk
from tkinter import messagebox

# ===========================================================
# CONFIGURAÇÕES
# ===========================================================
TIPOS_DOCUMENTO = [
    "extrato", "retificação", "alteração", "aditamento",
    "prorrogação", "termo de fomento", "termo de colaboração"
]

url_busca_negocios = 'https://diariooficial.prefeitura.sp.gov.br/md_epubli_controlador.php?acao=negocios_pesquisar'

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/148.0.0.0 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded',
}

# Caminhos do projeto
SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
PROJETO_DIR     = os.path.dirname(SCRIPT_DIR)
DATA_DIR        = os.path.join(PROJETO_DIR, "data")

ARQUIVO_ENTRADA = os.path.join(DATA_DIR, "Auditoria_Completa_Parcerias.xlsx")
CHECKPOINT_PATH = os.path.join(DATA_DIR, "checkpoint_negocios.json")
NOME_SAIDA      = os.path.join(DATA_DIR, "Auditoria_Completa_Parcerias_Negocios.xlsx")

os.makedirs(DATA_DIR, exist_ok=True)

# ===========================================================
# FUNÇÕES AUXILIARES
# ===========================================================
def formatar_processo(num):
    num = re.sub(r'\D', '', str(num))
    if len(num) == 16:
        return f"{num[:4]}.{num[4:8]}/{num[8:15]}-{num[15:]}"
    return num

def limpar_rotulo(texto):
    if not texto: return ""
    texto = str(texto).lower().replace(":", "")
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return " ".join(texto.split())

def extrair_informacoes_hibrido(html_content):
    soup  = BeautifulSoup(html_content, 'html.parser')
    linhas = list(soup.stripped_strings)

    modalidade  = "-"
    data_inicio = "-"
    data_fim    = "-"

    for i, linha in enumerate(linhas):
        linha_original = linha.strip()
        linha_limpa    = limpar_rotulo(linha_original)

        if "modalidade" in linha_limpa and modalidade == "-":
            if linha_limpa == "modalidade":
                if i + 1 < len(linhas): modalidade = linhas[i+1].strip()
            else:
                modalidade = linha_original.lower().replace("modalidade", "").replace(":", "").strip().title()

        elif "data de inicio" in linha_limpa and data_inicio == "-":
            if linha_limpa == "data de inicio":
                if i + 1 < len(linhas): data_inicio = linhas[i+1].strip()
            else:
                data_inicio = linha_original.split(":")[-1].strip() if ":" in linha_original else linha_original.replace("Data de Início", "").strip()

        elif "data de fim" in linha_limpa and data_fim == "-":
            if linha_limpa == "data de fim":
                if i + 1 < len(linhas): data_fim = linhas[i+1].strip()
            else:
                data_fim = linha_original.split(":")[-1].strip() if ":" in linha_original else linha_original.replace("Data de Fim", "").strip()

    return modalidade, data_inicio, data_fim

# ===========================================================
# FUNÇÕES DE CHECKPOINT
# ===========================================================
def salvar_checkpoint(indice, sucesso, sem_sucesso):
    dados = {
        "indice": indice,
        "sucesso": sucesso,
        "sem_sucesso": sem_sucesso,
    }
    with open(CHECKPOINT_PATH, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

def carregar_checkpoint():
    if not os.path.exists(CHECKPOINT_PATH):
        return None
    try:
        with open(CHECKPOINT_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def apagar_checkpoint():
    if os.path.exists(CHECKPOINT_PATH):
        os.remove(CHECKPOINT_PATH)

# ===========================================================
# 1. CARREGAMENTO DA PLANILHA ORIGINAL E FILTRAGEM
# ===========================================================
print("Iniciando o robô de Busca em Negócios...")

if not os.path.exists(ARQUIVO_ENTRADA):
    print(f"❌ Arquivo não encontrado: {ARQUIVO_ENTRADA}")
    print("Execute a extração principal primeiro!")
    exit()

try:
    df = pd.read_excel(ARQUIVO_ENTRADA)
except Exception as e:
    print(f"❌ Erro ao ler o arquivo: {e}")
    exit()

col_status = next((col for col in df.columns if "status" in col.lower()), None)
col_processo = next((col for col in df.columns if "processo" in col.lower()), None)

if not col_status or not col_processo:
    print("❌ Colunas de Status ou Processo não encontradas na planilha.")
    exit()

# Filtra apenas os que NÃO estão mapeados
indices_pendentes = df[df[col_status] != "Mapeado"].index.tolist()
total_pendentes = len(indices_pendentes)

print(f"📂 Arquivo de entrada: {ARQUIVO_ENTRADA}")
print(f"✅ {total_pendentes} processos NÃO MAPEADOS encontrados! Preparando motores...\n")

if total_pendentes == 0:
    print("Todos os processos já estão mapeados. Não há o que fazer!")
    exit()

# ===========================================================
# 2. VERIFICAR CHECKPOINT
# ===========================================================
checkpoint = carregar_checkpoint()

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)

recuperados      = 0
ainda_com_falha  = 0
indice_inicio    = 0

# Precisamos recarregar o df parcial caso exista um arquivo de saída em andamento
arquivo_parcial = ARQUIVO_ENTRADA.replace(".xlsx", "_Parcial_Negocios.xlsx")

if checkpoint and os.path.exists(arquivo_parcial):
    ja_feitos = checkpoint.get("indice", 0)
    resposta  = messagebox.askyesno(
        "Checkpoint encontrado",
        f"Foi encontrado um checkpoint com {ja_feitos} processos pesquisados na busca de negócios.\n\n"
        f"Deseja CONTINUAR de onde parou?\n\n"
        f"Clique 'Não' para começar do zero (o checkpoint será apagado).",
    )
    if resposta:
        recuperados      = checkpoint.get("sucesso", 0)
        ainda_com_falha  = checkpoint.get("sem_sucesso", 0)
        indice_inicio    = ja_feitos
        
        df = pd.read_excel(arquivo_parcial) # carrega os progressos já salvos
        print(f"▶️  Retomando do processo {indice_inicio + 1}/{total_pendentes} (já pesquisados: {indice_inicio})\n")
    else:
        apagar_checkpoint()
        if os.path.exists(arquivo_parcial): os.remove(arquivo_parcial)
        print("🔄 Iniciando do zero.\n")
else:
    apagar_checkpoint()
    if os.path.exists(arquivo_parcial): os.remove(arquivo_parcial)

# ===========================================================
# 3. EXECUÇÃO DA BUSCA EM NEGÓCIOS
# ===========================================================
print("-" * 60)
print(f"INICIANDO BUSCA DE {total_pendentes} PROCESSOS NO PAINEL DE NEGÓCIOS")
print("-" * 60)

tempo_inicio = time.time()
session = requests.Session()

for i in range(indice_inicio, total_pendentes):
    idx_df = indices_pendentes[i]
    processo_bruto = df.at[idx_df, col_processo]
    processo_fmt   = formatar_processo(processo_bruto)
    
    print(f"[{i+1}/{total_pendentes}] Buscando em negócios: {processo_fmt}...", end=" ")

    payload = {
        'hdnObjeto': processo_fmt,
        'hdnInicio': '0',
        'hdnModoPesquisa': 'DATA'
    }

    try:
        response = session.post(url_busca_negocios, headers=headers, data=payload, timeout=30)
    except requests.exceptions.RequestException as e:
        print(f"❌ Falha de conexão ({e})")
        ainda_com_falha += 1
        df.at[idx_df, col_status] = "Falha na conexão (Negócios)"
        df.to_excel(arquivo_parcial, index=False)
        salvar_checkpoint(i + 1, recuperados, ainda_com_falha)
        time.sleep(3)
        continue

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        
        documento_sucesso_link = None
        mod_final    = "-"
        dt_ini_final = "-"
        dt_fim_final = "-"

        links = soup.find_all('a', href=True)
        for tag_a in links:
            href = tag_a['href']
            
            if 'md_epubli_visualizar' in href:
                if not href.startswith('http'):
                    href = "https://diariooficial.prefeitura.sp.gov.br/" + href
                
                container = tag_a.find_parent(['div', 'p', 'tr', 'li'])
                texto_contexto = container.text.lower() if container else tag_a.parent.text.lower()
                
                if any(tipo in texto_contexto for tipo in TIPOS_DOCUMENTO):
                    try:
                        resp_doc = session.get(href, headers=headers, timeout=30)
                    except requests.exceptions.RequestException:
                        continue
                    
                    if 'iso-8859-1' in resp_doc.text.lower():
                        resp_doc.encoding = 'iso-8859-1'
                        
                    mod, d_ini, d_fim = extrair_informacoes_hibrido(resp_doc.text)
                    
                    if d_ini != "-" or d_fim != "-":
                        documento_sucesso_link = href
                        mod_final    = mod
                        dt_ini_final = d_ini
                        dt_fim_final = d_fim
                        break 
        
        if documento_sucesso_link:
            print("✅ Recuperado!")
            recuperados += 1
            df.at[idx_df, "Modalidade"]        = mod_final
            df.at[idx_df, "Data Início"]       = dt_ini_final
            df.at[idx_df, "Data Fim"]          = dt_fim_final
            df.at[idx_df, col_status]          = "Mapeado (Busca Negócios)"
            df.at[idx_df, "Link do Documento"] = documento_sucesso_link
        else:
            print("⚠️ Sem dados")
            ainda_com_falha += 1
            df.at[idx_df, col_status] = "Não mapeado (inclusive Negócios)"
    else:
        print("❌ Erro HTTP")
        ainda_com_falha += 1
        df.at[idx_df, col_status] = "Falha na busca (Negócios)"

    # Salva o arquivo parcial e checkpoint a cada rodada
    df.to_excel(arquivo_parcial, index=False)
    salvar_checkpoint(i + 1, recuperados, ainda_com_falha)
    time.sleep(0.3)

# ===========================================================
# 4. FINALIZAÇÃO
# ===========================================================
print("\n" + "=" * 60)
print("Gerando planilha Excel final...")

# Renomeia o parcial para o final
if os.path.exists(arquivo_parcial):
    if os.path.exists(NOME_SAIDA):
        os.remove(NOME_SAIDA)
    os.rename(arquivo_parcial, NOME_SAIDA)
else:
    df.to_excel(NOME_SAIDA, index=False)

apagar_checkpoint()

tempo_fim       = time.time()
tempo_total_seg = int(tempo_fim - tempo_inicio)
tempo_formatado = str(datetime.timedelta(seconds=tempo_total_seg))

print("=" * 60)
print("📊 INSIGHTS DA BUSCA EM NEGÓCIOS")
print("=" * 60)
print(f"-> {recuperados} processos foram RECUPERADOS com sucesso.")
print(f"-> {ainda_com_falha} processos continuam sem mapeamento.")
print(f"-> ⏱️  Tempo total de execução: {tempo_formatado}")
print(f"\n✅ Arquivo final salvo como: '{os.path.basename(NOME_SAIDA)}'")
print("🗑️  Checkpoint removido. Execução concluída!")
