"""
obs6.py — Observatório de Parcerias (versão Turbo: Session + Paralelo + Checkpoint)

Extrai dados detalhados de processos de parceria publicados no Diário Oficial de SP.
Recebe uma planilha com coluna 'codProcesso_y', busca cada processo no D.O. e extrai:
  - Órgão, Nome do Contratado, Número do Contrato
  - Objeto do Contrato, Modalidade
  - Data de Início / Data de Fim
  - Dotação Orçamentária, Natureza da Despesa
  - Link do Documento

Recursos:
  ✅ Checkpoint automático — retoma de onde parou após interrupção
  ✅ Backup automático a cada N processos (configurável)
  ✅ Execução paralela com múltiplos workers
  ✅ Retry automático em caso de falha de rede

Saídas (salvas em data/):
  - Auditoria_obs6_Parcerias.xlsx  ← planilha final
  - Backup_Temp_obs6.xlsx          ← backup intermediário
  - obs6_checkpoint.json           ← ponto de retomada

Uso:
  python src/obs6.py
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
import json
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- Caminhos ---
_DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
_DIR_RAIZ   = os.path.abspath(os.path.join(_DIR_SCRIPT, ".."))
_DIR_DATA   = os.path.join(_DIR_RAIZ, "data")

os.makedirs(_DIR_DATA, exist_ok=True)

# --- Configurações ---
CHECKPOINT_FILE  = os.path.join(_DIR_DATA, "obs6_checkpoint.json")
BACKUP_FILE      = os.path.join(_DIR_DATA, "Backup_Temp_obs6.xlsx")
NOME_ARQUIVO     = os.path.join(_DIR_DATA, "Auditoria_obs6_Parcerias.xlsx")
INTERVALO_BACKUP = 50       # salva a cada N processos concluídos
MAX_TENTATIVAS   = 3        # retries por processo
ESPERA_RETRY     = 20       # segundos entre retries
NUM_WORKERS      = 2        # requisições paralelas
SLEEP_ENTRE_REQ  = 0.5     # segundos de pausa por processo

TIPOS_DOCUMENTO = [
    "extrato", "retificacao", "alteracao", "aditamento",
    "prorrogacao", "termo de fomento", "termo de colaboracao"
]

url_busca = (
    'https://diariooficial.prefeitura.sp.gov.br/'
    'md_epubli_controlador.php?acao=materias_pesquisar'
)

base_headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/148.0.0.0 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded',
}

# --- Sessions por thread (reutiliza conexões TCP) ---
thread_local = threading.local()


def get_session():
    """Retorna (ou cria) uma Session dedicada para a thread atual."""
    if not hasattr(thread_local, 'session'):
        s = requests.Session()
        s.headers.update(base_headers)
        thread_local.session = s
    return thread_local.session


# --- Estado compartilhado ---
lock                    = threading.Lock()
stop_event              = threading.Event()
resultados_excel        = []
processos_com_sucesso   = 0
processos_sem_sucesso   = 0
processos_ja_feitos     = set()
contador_sessao         = 0


# --- 1. Seleção do arquivo ---
print("Iniciando obs6 (Versão Turbo: Session + Paralelo + Checkpoint)")
print("-" * 60)

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)

print("Selecione o arquivo com a coluna codProcesso_y...")

caminho_arquivo = filedialog.askopenfilename(
    title="Selecione o arquivo com a coluna codProcesso_y",
    filetypes=[
        ("Arquivos CSV", "*.csv"),
        ("Planilhas Excel", "*.xlsx"),
        ("Todos os arquivos", "*.*")
    ]
)

if not caminho_arquivo:
    print("Nenhum arquivo selecionado. Encerrando.")
    exit()

try:
    if caminho_arquivo.lower().endswith(('.xlsx', '.xls')):
        df_entrada = pd.read_excel(caminho_arquivo)
    else:
        df_entrada = pd.read_csv(caminho_arquivo, sep=None, engine='python', encoding='latin-1')

    PROCESSOS = df_entrada['codProcesso_y'].dropna().astype(str).tolist()
    print(f"[OK] {len(PROCESSOS)} processos carregados!\n")

except Exception as e:
    print(f"[ERRO] {e}")
    exit()


# --- 2. Checkpoint ---
if os.path.exists(CHECKPOINT_FILE):
    try:
        with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
            checkpoint_lido = json.load(f)
        arquivo_anterior = checkpoint_lido.get("arquivo_entrada", "")
    except Exception:
        checkpoint_lido = {}
        arquivo_anterior = ""

    if arquivo_anterior and os.path.normpath(arquivo_anterior) != os.path.normpath(caminho_arquivo):
        aviso = (
            f"ATENÇÃO: Checkpoint de arquivo DIFERENTE!\n\n"
            f"Anterior: {os.path.basename(arquivo_anterior)}\n"
            f"Atual:    {os.path.basename(caminho_arquivo)}\n\n"
            f"O script vai começar do ZERO."
        )
        messagebox.showwarning("Arquivo diferente", aviso, parent=root)
        os.remove(CHECKPOINT_FILE)
        if os.path.exists(BACKUP_FILE):
            os.remove(BACKUP_FILE)
        print("[AVISO] Checkpoint ignorado. Iniciando do zero.\n")
    else:
        resposta = messagebox.askyesno(
            "Retomada detectada",
            f"Checkpoint encontrado.\nArquivo: {os.path.basename(caminho_arquivo)}\n\n"
            f"Deseja CONTINUAR de onde parou?\n(Não = começar do zero)",
            parent=root
        )
        if resposta:
            try:
                processos_ja_feitos   = set(checkpoint_lido.get("processos_feitos", []))
                processos_com_sucesso = checkpoint_lido.get("processos_com_sucesso", 0)
                processos_sem_sucesso = checkpoint_lido.get("processos_sem_sucesso", 0)
                if os.path.exists(BACKUP_FILE):
                    resultados_excel = pd.read_excel(BACKUP_FILE).to_dict('records')
                print(f"[RETOMADA] Pulando {len(processos_ja_feitos)} processos já concluídos.")
                print(f"[RETOMADA] Retomando a partir do processo {len(processos_ja_feitos) + 1}...\n")
            except Exception as e:
                print(f"[AVISO] Falha ao ler checkpoint: {e}. Iniciando do zero.")
                processos_ja_feitos = set()
                resultados_excel    = []
        else:
            os.remove(CHECKPOINT_FILE)
            if os.path.exists(BACKUP_FILE):
                os.remove(BACKUP_FILE)
            print("[INFO] Iniciando do zero.\n")


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
    soup = BeautifulSoup(html_content, 'html.parser')
    linhas = list(soup.stripped_strings)
    modalidade = data_inicio = data_fim = orgao = numero_contrato = objeto = \
        nome_contratado = dotacao = natureza_despesa = "-"

    for i, linha in enumerate(linhas):
        linha_original = linha.strip()
        linha_limpa = limpar_rotulo(linha_original)

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
                time.sleep(ESPERA_RETRY)
            else:
                raise e


def salvar_checkpoint():
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "arquivo_entrada": caminho_arquivo,
            "processos_feitos": list(processos_ja_feitos),
            "processos_com_sucesso": processos_com_sucesso,
            "processos_sem_sucesso": processos_sem_sucesso
        }, f, ensure_ascii=False)


def processar_um_processo(args):
    """Função executada por cada worker thread."""
    processo_bruto, pos_global, total = args

    if stop_event.is_set():
        return None

    session = get_session()
    processo_fmt = formatar_processo(processo_bruto)
    processo_codificado = urllib.parse.quote(processo_fmt)
    payload = (
        f"hdnTermoPesquisa={processo_codificado}&hdnTipoPesquisa=Q&hdnVersaoDiario="
        f"&hdnOndePesquisa=&hdnTipoDataPesquisa=I&hdnDataInicioPesquisa="
        f"&hdnDataFimPesquisa=&hdnTipoDocumentoPesquisa=&hdnVeiculoPublicacao="
        f"&hdnDataPublicacao=&hdnOrgaoFiltro=&hdnUnidadeResponsavelFiltro="
        f"&hdnTipoDocumentoFiltro=&hdnInicio=0&hdnVisualizacao=L&hdnModoPesquisa=RAPIDA"
    )

    try:
        response = fazer_requisicao_com_retry(
            lambda: session.post(url_busca, data=payload, timeout=60)
        )
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        documento_sucesso_link = None
        mod_f = dt_ini_f = dt_fim_f = org_f = num_c_f = obj_f = nome_c_f = dot_f = nat_f = "-"

        for tag_a in soup.find_all('a', href=True):
            if stop_event.is_set():
                break
            href = tag_a['href']
            if 'md_epubli' in href or 'materia' in href or 'acao=exibir' in href:
                if "pesquisar" not in href and "login" not in href:
                    if not href.startswith('http'):
                        href = "https://diariooficial.prefeitura.sp.gov.br/" + href

                    container = tag_a.find_parent(['div', 'p', 'tr', 'li'])
                    texto_contexto = container.text.lower() if container else tag_a.parent.text.lower()

                    if any(tipo in texto_contexto for tipo in TIPOS_DOCUMENTO):
                        try:
                            resp_doc = fazer_requisicao_com_retry(
                                lambda h=href: session.get(h, timeout=60)
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

        time.sleep(SLEEP_ENTRE_REQ)

        if documento_sucesso_link:
            print(f"[{pos_global}/{total}] {processo_fmt} [OK]")
            return ("sucesso", processo_fmt, {
                "Nr do Processo": processo_fmt,
                "Orgao": org_f, "Nome do Contratado": nome_c_f, "Nr do Contrato": num_c_f,
                "Objeto do Contrato": obj_f, "Modalidade": mod_f,
                "Data Inicio": dt_ini_f, "Data Fim": dt_fim_f,
                "Dotacao Orcamentaria": dot_f, "Natureza da Despesa": nat_f,
                "Status": "Mapeado", "Link do Documento": documento_sucesso_link
            })
        else:
            print(f"[{pos_global}/{total}] {processo_fmt} [SEM DATAS]")
            return ("sem_dados", processo_fmt, {
                "Nr do Processo": processo_fmt,
                "Orgao": "-", "Nome do Contratado": "-", "Nr do Contrato": "-",
                "Objeto do Contrato": "-", "Modalidade": "-",
                "Data Inicio": "-", "Data Fim": "-",
                "Dotacao Orcamentaria": "-", "Natureza da Despesa": "-",
                "Status": "Nao mapeado (sem datas)", "Link do Documento": "-"
            })

    except requests.exceptions.RequestException:
        print(f"[{pos_global}/{total}] {processo_fmt} [ERRO CONEXAO]")
        return ("erro", processo_fmt, {
            "Nr do Processo": processo_fmt,
            "Orgao": "-", "Nome do Contratado": "-", "Nr do Contrato": "-",
            "Objeto do Contrato": "-", "Modalidade": "-",
            "Data Inicio": "-", "Data Fim": "-",
            "Dotacao Orcamentaria": "-", "Natureza da Despesa": "-",
            "Status": "Falha na conexao (todas as tentativas)", "Link do Documento": "-"
        })


# --- Execução em lote (paralela) ---
total     = len(PROCESSOS)
pendentes = [p for p in PROCESSOS if formatar_processo(p) not in processos_ja_feitos]

if len(processos_ja_feitos) > 0:
    overlap = total - len(pendentes)
    if overlap < len(processos_ja_feitos) * 0.5:
        aviso = (
            f"ATENÇÃO: O checkpoint parece estar corrompido!\n\n"
            f"Processos no checkpoint: {len(processos_ja_feitos)}\n"
            f"Processos reconhecidos no arquivo atual: {overlap}\n\n"
            f"Deseja IGNORAR o checkpoint e começar do zero?"
        )
        if messagebox.askyesno("Checkpoint suspeito", aviso, parent=root):
            processos_ja_feitos.clear()
            resultados_excel.clear()
            processos_com_sucesso = 0
            processos_sem_sucesso = 0
            pendentes = list(PROCESSOS)
            if os.path.exists(CHECKPOINT_FILE): os.remove(CHECKPOINT_FILE)
            if os.path.exists(BACKUP_FILE):     os.remove(BACKUP_FILE)
            print("[INFO] Checkpoint ignorado. Iniciando do zero.\n")

print("-" * 60)
print(f"TOTAL: {total} | JÁ FEITOS: {len(processos_ja_feitos)} | PENDENTES: {len(pendentes)}")
print(f"Workers paralelos: {NUM_WORKERS} | Sleep: {SLEEP_ENTRE_REQ}s | Retries: {MAX_TENTATIVAS}x")
print("-" * 60)

inicio_ja_feitos = len(processos_ja_feitos)
tempo_inicio     = time.time()

args_list = [
    (p, inicio_ja_feitos + i + 1, total)
    for i, p in enumerate(pendentes)
]

try:
    with ThreadPoolExecutor(max_workers=NUM_WORKERS) as executor:
        futures = {executor.submit(processar_um_processo, args): args for args in args_list}

        for future in as_completed(futures):
            if stop_event.is_set():
                break

            resultado = future.result()
            if resultado is None:
                continue

            status_tipo, processo_fmt, dados = resultado

            with lock:
                resultados_excel.append(dados)
                processos_ja_feitos.add(processo_fmt)

                if status_tipo == "sucesso":
                    processos_com_sucesso += 1
                else:
                    processos_sem_sucesso += 1

                contador_sessao += 1

                if contador_sessao % INTERVALO_BACKUP == 0:
                    try:
                        pd.DataFrame(resultados_excel).to_excel(BACKUP_FILE, index=False)
                        salvar_checkpoint()
                        print(f"   [BACKUP + CHECKPOINT] {contador_sessao}/{len(pendentes)} "
                              "processados nesta sessão")
                    except PermissionError:
                        print(f"   [AVISO] Backup não salvo ({os.path.basename(BACKUP_FILE)} "
                              "está aberto). Feche o arquivo e o backup será retomado.")

except KeyboardInterrupt:
    print("\n\n[INTERRUPÇÃO] Ctrl+C detectado! Salvando progresso...")
    stop_event.set()
    with lock:
        pd.DataFrame(resultados_excel).to_excel(BACKUP_FILE, index=False)
        salvar_checkpoint()
    print(f"[OK] Backup salvo: {os.path.basename(BACKUP_FILE)}")
    print(f"[OK] Checkpoint salvo: {os.path.basename(CHECKPOINT_FILE)}")
    print(f"[INFO] {len(processos_ja_feitos)}/{total} processos concluídos.")
    print("[INFO] Rode o obs6.py novamente e escolha 'Continuar' para retomar.")
    exit()

# --- Salvamento final ---
tempo_fim        = time.time()
tempo_formatado  = str(datetime.timedelta(seconds=int(tempo_fim - tempo_inicio)))

print("\n" + "=" * 60)
print("Gerando planilha Excel final...")
df_resultados = pd.DataFrame(resultados_excel)
try:
    df_resultados.to_excel(NOME_ARQUIVO, index=False)
except PermissionError:
    NOME_ARQUIVO_ALT = NOME_ARQUIVO.replace(".xlsx", "_v2.xlsx")
    print(f"[AVISO] Arquivo aberto. Salvando como '{os.path.basename(NOME_ARQUIVO_ALT)}'...")
    df_resultados.to_excel(NOME_ARQUIVO_ALT, index=False)
    NOME_ARQUIVO = NOME_ARQUIVO_ALT

if os.path.exists(CHECKPOINT_FILE):
    os.remove(CHECKPOINT_FILE)

print("=" * 60)
print("INSIGHTS DA EXTRAÇÃO (OBS6 TURBO)")
print("=" * 60)
print(f"-> {processos_com_sucesso} processos mapeados com sucesso.")
print(f"-> {processos_sem_sucesso} processos não mapeados / com falha.")
print(f"-> Tempo desta sessão: {tempo_formatado}")
print(f"\n[CONCLUÍDO] Arquivo salvo em: data/{os.path.basename(NOME_ARQUIVO)}")
