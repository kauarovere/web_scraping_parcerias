"""
obs6_retry.py — Reprocessador de falhas de conexão do obs6.

Quando o obs6.py termina com processos em status 'Falha na conexão',
execute este script para tentar novamente apenas os que falharam.
Os resultados são gravados de volta na planilha original com status atualizado.

Uso:
  python src/obs6_retry.py
  (Selecione a planilha gerada pelo obs6 — Auditoria_obs6_Parcerias.xlsx ou Backup)
"""

import requests
from bs4 import BeautifulSoup
import re
import urllib.parse
import unicodedata
import pandas as pd
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import datetime
import os

# --- Configurações ---
MAX_TENTATIVAS = 3
ESPERA_RETRY   = 30   # segundos entre retries
STATUS_FALHA   = "Falha na conexao (todas as tentativas)"

TIPOS_DOCUMENTO = [
    "extrato", "retificacao", "alteracao", "aditamento",
    "prorrogacao", "termo de fomento", "termo de colaboracao"
]

url_busca = (
    'https://diariooficial.prefeitura.sp.gov.br/'
    'md_epubli_controlador.php?acao=materias_pesquisar'
)

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/148.0.0.0 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded',
}


# --- 1. Selecionar planilha gerada pelo obs6 ---
print("obs6_retry — Reprocessador de falhas de conexão")
print("-" * 60)

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)

print("Selecione a planilha gerada pelo obs6 (Auditoria_obs6_Parcerias.xlsx ou Backup)...")

caminho_planilha = filedialog.askopenfilename(
    title="Selecione a planilha do obs6 com os erros de conexão",
    filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
)

if not caminho_planilha:
    print("Nenhum arquivo selecionado. Encerrando.")
    exit()


# --- 2. Ler planilha e filtrar falhas ---
try:
    df = pd.read_excel(caminho_planilha)
except Exception as e:
    print(f"[ERRO] Não foi possível ler a planilha: {e}")
    exit()

col_status = next((col for col in df.columns if "status" in col.lower()), None)
if col_status is None:
    print("[ERRO] Coluna 'Status' não encontrada na planilha.")
    exit()

col_processo = next((col for col in df.columns if "processo" in col.lower()), None)
if col_processo is None:
    print("[ERRO] Coluna de processo não encontrada na planilha.")
    exit()

df_falhas    = df[df[col_status] == STATUS_FALHA].copy()
total_falhas = len(df_falhas)

if total_falhas == 0:
    messagebox.showinfo(
        "Sem falhas",
        "Nenhum processo com erro de conexão encontrado na planilha!\n\nNada a reprocessar.",
        parent=root
    )
    print("[INFO] Nenhum processo com erro de conexão encontrado. Encerrando.")
    exit()

confirmar = messagebox.askyesno(
    "Reprocessar falhas",
    f"Foram encontrados {total_falhas} processos com erro de conexão.\n\n"
    f"Deseja reprocessá-los agora?\n\n"
    f"Os resultados serão atualizados na planilha original.",
    parent=root
)

if not confirmar:
    print("Operação cancelada pelo usuário.")
    exit()

print(f"\n[OK] {total_falhas} processos com falha encontrados. Iniciando reprocessamento...\n")
print("-" * 60)


# --- Funções auxiliares ---

def formatar_processo(num):
    num = re.sub(r'\D', '', str(num))
    if len(num) == 16:
        return f"{num[:4]}.{num[4:8]}/{num[8:15]}-{num[15:]}"
    return num


def limpar_rotulo(texto):
    if not texto:
        return ""
    texto = str(texto).lower().replace(":", "")
    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )
    return " ".join(texto.split())


def extrair_informacoes_hibrido(html_content):
    """Extrai campos estruturados de um documento HTML do Diário Oficial."""
    soup  = BeautifulSoup(html_content, 'html.parser')
    linhas = list(soup.stripped_strings)
    modalidade = data_inicio = data_fim = orgao = numero_contrato = objeto = \
        nome_contratado = dotacao = natureza_despesa = "-"

    for i, linha in enumerate(linhas):
        linha_original = linha.strip()
        linha_limpa    = limpar_rotulo(linha_original)

        if "modalidade" in linha_limpa and modalidade == "-":
            if linha_limpa == "modalidade":
                if i + 1 < len(linhas): modalidade = linhas[i+1].strip()
            else:
                modalidade = re.sub(r'(?i)modalidade\s*:?', '', linha_original).strip().title()
        elif "data de inicio" in linha_limpa and data_inicio == "-":
            if linha_limpa == "data de inicio":
                if i + 1 < len(linhas): data_inicio = linhas[i+1].strip()
            else:
                data_inicio = re.sub(r'(?i)data de in[íi]cio\s*:?', '', linha_original).strip()
        elif "data de fim" in linha_limpa and data_fim == "-":
            if linha_limpa == "data de fim":
                if i + 1 < len(linhas): data_fim = linhas[i+1].strip()
            else:
                data_fim = re.sub(r'(?i)data de fim\s*:?', '', linha_original).strip()
        elif "orgao" in linha_limpa and orgao == "-":
            if linha_limpa == "orgao":
                if i + 1 < len(linhas): orgao = linhas[i+1].strip()
            else:
                orgao = re.sub(r'(?i)[óo]rg[ãa]o\s*:?', '', linha_original).strip()
        elif "numero do contrato" in linha_limpa and numero_contrato == "-":
            if linha_limpa == "numero do contrato":
                if i + 1 < len(linhas): numero_contrato = linhas[i+1].strip()
            else:
                numero_contrato = re.sub(r'(?i)n[úu]mero do contrato\s*:?', '', linha_original).strip()
        elif "objeto do contrato" in linha_limpa and objeto == "-":
            if linha_limpa == "objeto do contrato":
                if i + 1 < len(linhas): objeto = linhas[i+1].strip()
            else:
                objeto = re.sub(r'(?i)objeto do contrato\s*:?', '', linha_original).strip()
        elif "nome do contratado" in linha_limpa and nome_contratado == "-":
            if "nome do contratado" in linha_limpa:
                if i + 1 < len(linhas): nome_contratado = linhas[i+1].strip()
            else:
                nome_contratado = re.sub(
                    r'(?i)nome do contratado\s*(\(entidade parceira\))?\s*:?', '',
                    linha_original
                ).strip()
        elif "dotacao orcamentaria" in linha_limpa and dotacao == "-":
            if linha_limpa == "dotacao orcamentaria":
                if i + 1 < len(linhas): dotacao = linhas[i+1].strip()
            else:
                dotacao = re.sub(r'(?i)dota[çc][ãa]o or[çc]ament[áa]ria\s*:?', '', linha_original).strip()
        elif "natureza da despesa" in linha_limpa and natureza_despesa == "-":
            if linha_limpa == "natureza da despesa":
                if i + 1 < len(linhas): natureza_despesa = linhas[i+1].strip()
            else:
                natureza_despesa = re.sub(r'(?i)natureza da despesa\s*:?', '', linha_original).strip()

    return (modalidade, data_inicio, data_fim, orgao, numero_contrato,
            objeto, nome_contratado, dotacao, natureza_despesa)


def fazer_requisicao_com_retry(func_request):
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            return func_request()
        except requests.exceptions.RequestException as e:
            if tentativa < MAX_TENTATIVAS:
                print(f"\n   [RETRY {tentativa}/{MAX_TENTATIVAS}] Aguardando {ESPERA_RETRY}s...", end=" ")
                time.sleep(ESPERA_RETRY)
            else:
                raise e


# --- 3. Reprocessar os processos com falha ---
recuperados      = 0
ainda_com_falha  = 0
tempo_inicio     = time.time()
indices_falha    = df_falhas.index.tolist()

for i, idx in enumerate(indices_falha):
    processo_bruto  = str(df.at[idx, col_processo])
    processo_fmt    = formatar_processo(processo_bruto)
    print(f"[{i+1}/{total_falhas}] Reprocessando: {processo_fmt}...", end=" ")

    processo_codificado = urllib.parse.quote(processo_fmt)
    payload_bruto = (
        f"hdnTermoPesquisa={processo_codificado}&hdnTipoPesquisa=Q&hdnVersaoDiario="
        f"&hdnOndePesquisa=&hdnTipoDataPesquisa=I&hdnDataInicioPesquisa="
        f"&hdnDataFimPesquisa=&hdnTipoDocumentoPesquisa=&hdnVeiculoPublicacao="
        f"&hdnDataPublicacao=&hdnOrgaoFiltro=&hdnUnidadeResponsavelFiltro="
        f"&hdnTipoDocumentoFiltro=&hdnInicio=0&hdnVisualizacao=L&hdnModoPesquisa=RAPIDA"
    )

    try:
        response = fazer_requisicao_com_retry(
            lambda: requests.post(url_busca, headers=headers, data=payload_bruto, timeout=60)
        )
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        documento_sucesso_link = None
        mod_f = dt_ini_f = dt_fim_f = org_f = num_c_f = obj_f = nome_c_f = dot_f = nat_f = "-"

        for tag_a in soup.find_all('a', href=True):
            href = tag_a['href']
            if 'md_epubli' in href or 'materia' in href or 'acao=exibir' in href:
                if "pesquisar" not in href and "login" not in href:
                    if not href.startswith('http'):
                        href = "https://diariooficial.prefeitura.sp.gov.br/" + href

                    container    = tag_a.find_parent(['div', 'p', 'tr', 'li'])
                    texto_contexto = container.text.lower() if container else tag_a.parent.text.lower()

                    if any(tipo in texto_contexto for tipo in TIPOS_DOCUMENTO):
                        try:
                            resp_doc = fazer_requisicao_com_retry(
                                lambda h=href: requests.get(h, headers=headers, timeout=60)
                            )
                            if 'iso-8859-1' in resp_doc.text.lower():
                                resp_doc.encoding = 'iso-8859-1'

                            mod, d_ini, d_fim, org, num_c, obj, nome_c, dot, nat = \
                                extrair_informacoes_hibrido(resp_doc.text)

                            if d_ini != "-" or d_fim != "-":
                                documento_sucesso_link = href
                                mod_f, dt_ini_f, dt_fim_f = mod, d_ini, d_fim
                                org_f, num_c_f, obj_f     = org, num_c, obj
                                nome_c_f, dot_f, nat_f    = nome_c, dot, nat
                                break
                        except requests.exceptions.RequestException:
                            continue

        if documento_sucesso_link:
            print("[OK] Recuperado!")
            recuperados += 1
            df.at[idx, "Orgao"]                = org_f
            df.at[idx, "Nome do Contratado"]   = nome_c_f
            df.at[idx, "Nr do Contrato"]        = num_c_f
            df.at[idx, "Objeto do Contrato"]   = obj_f
            df.at[idx, "Modalidade"]           = mod_f
            df.at[idx, "Data Inicio"]           = dt_ini_f
            df.at[idx, "Data Fim"]             = dt_fim_f
            df.at[idx, "Dotacao Orcamentaria"] = dot_f
            df.at[idx, "Natureza da Despesa"]  = nat_f
            df.at[idx, col_status]             = "Mapeado (retry)"
            df.at[idx, "Link do Documento"]    = documento_sucesso_link
        else:
            print("[SEM DATAS]")
            ainda_com_falha += 1
            df.at[idx, col_status] = "Nao mapeado (sem datas no retry)"

    except requests.exceptions.RequestException:
        print("[ERRO] Ainda com falha de conexão.")
        ainda_com_falha += 1
        df.at[idx, col_status] = "Falha persistente na conexao"

    time.sleep(1)


# --- 4. Salvar planilha atualizada ---
tempo_fim       = time.time()
tempo_formatado = str(datetime.timedelta(seconds=int(tempo_fim - tempo_inicio)))

nome_saida = caminho_planilha.replace(".xlsx", "_atualizado.xlsx")
df.to_excel(nome_saida, index=False)

print("\n" + "=" * 60)
print("RESULTADO DO REPROCESSAMENTO")
print("=" * 60)
print(f"-> {recuperados} processos RECUPERADOS com sucesso.")
print(f"-> {ainda_com_falha} processos ainda com falha.")
print(f"-> Tempo de execução: {tempo_formatado}")
print(f"\n[CONCLUÍDO] Planilha atualizada salva como: '{os.path.basename(nome_saida)}'")
