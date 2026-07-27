import requests
from bs4 import BeautifulSoup
import re
import urllib.parse
import unicodedata
import pandas as pd
import time
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# ===========================================================
# CONFIGURAÇÕES
# ===========================================================
TIPOS_DOCUMENTO = [
    "extrato", "retificação", "alteração", "aditamento",
    "prorrogação", "termo de fomento", "termo de colaboração"
]

url_busca = 'https://diariooficial.prefeitura.sp.gov.br/md_epubli_controlador.php?acao=materias_pesquisar'

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/148.0.0.0 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded',
}

# Caminho do arquivo de checkpoint (salvo na pasta data/ do projeto)
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
PROJETO_DIR  = os.path.dirname(SCRIPT_DIR)
DATA_DIR     = os.path.join(PROJETO_DIR, "data")
CHECKPOINT_PATH = os.path.join(DATA_DIR, "checkpoint_extracao.json")
NOME_SAIDA   = os.path.join(DATA_DIR, "Auditoria_Completa_Parcerias.xlsx")

os.makedirs(DATA_DIR, exist_ok=True)

# ===========================================================
# FUNÇÕES AUXILIARES
# ===========================================================

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

# ===========================================================
# FUNÇÕES DE CHECKPOINT
# ===========================================================

def salvar_checkpoint(arquivo_entrada, processos_feitos, resultados, sucesso, sem_sucesso):
    """Persiste o estado atual no arquivo de checkpoint."""
    dados = {
        "arquivo_entrada":  arquivo_entrada,
        "processos_feitos": processos_feitos,
        "resultados":       resultados,
        "sucesso":          sucesso,
        "sem_sucesso":      sem_sucesso,
    }
    with open(CHECKPOINT_PATH, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

def carregar_checkpoint():
    """Lê o checkpoint salvo. Retorna None se não existir."""
    if not os.path.exists(CHECKPOINT_PATH):
        return None
    try:
        with open(CHECKPOINT_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def apagar_checkpoint():
    """Remove o checkpoint ao final de uma execução completa."""
    if os.path.exists(CHECKPOINT_PATH):
        os.remove(CHECKPOINT_PATH)

# ===========================================================
# 1. SELEÇÃO DO ARQUIVO COM O MOUSE
# ===========================================================
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

# ===========================================================
# 2. LEITURA DO ARQUIVO DE ENTRADA
# ===========================================================
try:
    if caminho_arquivo.lower().endswith(('.xlsx', '.xls')):
        df_entrada = pd.read_excel(caminho_arquivo)
    else:
        df_entrada = pd.read_csv(caminho_arquivo, sep=None, engine='python', encoding='latin-1')

    PROCESSOS = df_entrada['codProcesso_y'].dropna().astype(str).tolist()
    print(f"✅ {len(PROCESSOS)} processos carregados com sucesso! Preparando motores...\n")

except Exception as e:
    print(f"❌ Erro ao ler o arquivo selecionado: {e}")
    print("Dica: Se a coluna tiver outro nome na planilha atual, o script não vai encontrá-la.")
    exit()

# ===========================================================
# 3. VERIFICAR CHECKPOINT — RETOMAR OU COMEÇAR DO ZERO?
# ===========================================================
checkpoint = carregar_checkpoint()

resultados_excel       = []
processos_com_sucesso  = 0
processos_sem_sucesso  = 0
indice_inicio          = 0  # de onde o loop vai começar

if checkpoint and checkpoint.get("arquivo_entrada") == caminho_arquivo:
    ja_feitos = checkpoint.get("processos_feitos", 0)
    resposta  = messagebox.askyesno(
        "Checkpoint encontrado",
        f"Foi encontrado um checkpoint com {ja_feitos} processos já concluídos.\n\n"
        f"Deseja CONTINUAR de onde parou?\n\n"
        f"Clique 'Não' para começar do zero (o checkpoint será apagado).",
        parent=root
    )
    if resposta:
        resultados_excel      = checkpoint.get("resultados", [])
        processos_com_sucesso = checkpoint.get("sucesso", 0)
        processos_sem_sucesso = checkpoint.get("sem_sucesso", 0)
        indice_inicio         = ja_feitos
        print(f"▶️  Retomando do processo {indice_inicio + 1}/{len(PROCESSOS)} (já feitos: {indice_inicio})\n")
    else:
        apagar_checkpoint()
        print("🔄 Iniciando do zero.\n")
else:
    # Checkpoint de outro arquivo ou não existe — ignora
    if checkpoint:
        print("⚠️  Checkpoint encontrado é de outro arquivo. Iniciando do zero.\n")
        apagar_checkpoint()

# ===========================================================
# 4. EXECUÇÃO EM LOTE (COM CHECKPOINT AUTOMÁTICO)
# ===========================================================
print("-" * 60)
print(f"INICIANDO EXTRAÇÃO DE {len(PROCESSOS)} PROCESSOS (BASE COMPLETA)")
print("-" * 60)

for i in range(indice_inicio, len(PROCESSOS)):
    processo_bruto = PROCESSOS[i]
    processo_fmt   = formatar_processo(processo_bruto)
    print(f"[{i+1}/{len(PROCESSOS)}] Buscando processo: {processo_fmt}...", end=" ")

    processo_codificado = urllib.parse.quote(processo_fmt)
    payload_bruto = (
        f"hdnTermoPesquisa={processo_codificado}&hdnTipoPesquisa=Q&hdnVersaoDiario="
        f"&hdnOndePesquisa=&hdnTipoDataPesquisa=I&hdnDataInicioPesquisa="
        f"&hdnDataFimPesquisa=&hdnTipoDocumentoPesquisa=&hdnVeiculoPublicacao="
        f"&hdnDataPublicacao=&hdnOrgaoFiltro=&hdnUnidadeResponsavelFiltro="
        f"&hdnTipoDocumentoFiltro=&hdnInicio=0&hdnVisualizacao=L&hdnModoPesquisa=RAPIDA"
    )

    try:
        response = requests.post(url_busca, headers=headers, data=payload_bruto, timeout=30)
    except requests.exceptions.RequestException as e:
        print(f"❌ Falha de conexão ({e})")
        processos_sem_sucesso += 1
        resultados_excel.append({
            "Nº do Processo":  processo_fmt,
            "Modalidade":      "Erro de Conexão",
            "Data Início":     "-",
            "Data Fim":        "-",
            "Status":          "Falha na conexão",
            "Link do Documento": "-"
        })
        # Salva checkpoint mesmo no erro para não perder o progresso
        salvar_checkpoint(caminho_arquivo, i + 1, resultados_excel,
                          processos_com_sucesso, processos_sem_sucesso)
        time.sleep(3)  # Espera um pouco antes de continuar após falha
        continue

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')

        documento_sucesso_link = None
        mod_final    = "Não identificada"
        dt_ini_final = "Não identificada"
        dt_fim_final = "Não identificada"

        for tag_a in soup.find_all('a', href=True):
            href = tag_a['href']

            if 'md_epubli' in href or 'materia' in href or 'acao=exibir' in href:
                if "pesquisar" not in href and "login" not in href:
                    if not href.startswith('http'):
                        href = "https://diariooficial.prefeitura.sp.gov.br/" + href

                    container      = tag_a.find_parent(['div', 'p', 'tr', 'li'])
                    texto_contexto = container.text.lower() if container else tag_a.parent.text.lower()

                    if any(tipo in texto_contexto for tipo in TIPOS_DOCUMENTO):
                        try:
                            resp_doc = requests.get(href, headers=headers, timeout=30)
                        except requests.exceptions.RequestException:
                            continue

                        if 'iso-8859-1' in resp_doc.text.lower():
                            resp_doc.encoding = 'iso-8859-1'

                        mod, d_ini, d_fim = extrair_informacoes_hibrido(resp_doc.text)

                        if d_ini != "Não identificada" or d_fim != "Não identificada":
                            documento_sucesso_link = href
                            mod_final    = mod
                            dt_ini_final = d_ini
                            dt_fim_final = d_fim
                            break

        if documento_sucesso_link:
            print("✅ OK")
            processos_com_sucesso += 1
            resultados_excel.append({
                "Nº do Processo":   processo_fmt,
                "Modalidade":       mod_final,
                "Data Início":      dt_ini_final,
                "Data Fim":         dt_fim_final,
                "Status":           "Mapeado",
                "Link do Documento": documento_sucesso_link
            })
        else:
            print("⚠️ Sem dados")
            processos_sem_sucesso += 1
            resultados_excel.append({
                "Nº do Processo":   processo_fmt,
                "Modalidade":       "-",
                "Data Início":      "-",
                "Data Fim":         "-",
                "Status":           "Não mapeado (Rótulos ausentes)",
                "Link do Documento": "-"
            })
    else:
        print("❌ Erro HTTP")
        processos_sem_sucesso += 1
        resultados_excel.append({
            "Nº do Processo":   processo_fmt,
            "Modalidade":       "Erro de Conexão",
            "Data Início":      "-",
            "Data Fim":         "-",
            "Status":           "Falha na busca",
            "Link do Documento": "-"
        })

    # --- SALVA CHECKPOINT APÓS CADA PROCESSO ---
    salvar_checkpoint(caminho_arquivo, i + 1, resultados_excel,
                      processos_com_sucesso, processos_sem_sucesso)

    time.sleep(1)  # Respiro de 1 segundo para o firewall não nos derrubar

# ===========================================================
# 5. GERAÇÃO DO ARQUIVO FINAL
# ===========================================================
print("\n" + "=" * 60)
print("Gerando planilha Excel final...")

df_resultados = pd.DataFrame(resultados_excel)
df_resultados.to_excel(NOME_SAIDA, index=False)

# Checkpoint concluído — pode apagar com segurança
apagar_checkpoint()

# --- INSIGHTS FINAIS ---
print("=" * 60)
print("📊 INSIGHTS DA EXTRAÇÃO COMPLETA")
print("=" * 60)
print(f"-> {processos_com_sucesso} processos foram mapeados com sucesso.")
print(f"-> {processos_sem_sucesso} processos não foram mapeados com sucesso.")
print(f"\n✅ Arquivo final salvo como: '{NOME_SAIDA}'")
print("🗑️  Checkpoint removido. Execução concluída!")