"""
obs6_criar_checkpoint.py — Utilitário para criar checkpoint a partir de um backup.

Use quando o checkpoint foi perdido mas você ainda tem o arquivo de backup
(Backup_Temp_obs6.xlsx). Este script reconstrói o checkpoint para que o
obs6.py possa retomar de onde parou.

Uso:
  python src/obs6_criar_checkpoint.py
  (Selecione o arquivo Backup_Temp_obs6.xlsx)
"""

import pandas as pd
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Caminhos ---
_DIR_SCRIPT     = os.path.dirname(os.path.abspath(__file__))
_DIR_RAIZ       = os.path.abspath(os.path.join(_DIR_SCRIPT, ".."))
_DIR_DATA       = os.path.join(_DIR_RAIZ, "data")
CHECKPOINT_FILE = os.path.join(_DIR_DATA, "obs6_checkpoint.json")

os.makedirs(_DIR_DATA, exist_ok=True)

print("obs6_criar_checkpoint — Conversor de backup para checkpoint")
print("-" * 60)

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)

# --- Selecionar o backup ---
print("Selecione o arquivo de BACKUP do obs6 (Backup_Temp_obs6.xlsx)...")

caminho_backup = filedialog.askopenfilename(
    title="Selecione o Backup_Temp_obs6.xlsx",
    filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
)

if not caminho_backup:
    print("Nenhum arquivo selecionado. Encerrando.")
    exit()

# --- Ler o backup ---
try:
    df_backup = pd.read_excel(caminho_backup)
except Exception as e:
    print(f"[ERRO] Não foi possível ler o backup: {e}")
    exit()

# Identifica coluna do processo automaticamente
col_processo = next((col for col in df_backup.columns if "processo" in col.lower()), None)
if col_processo is None:
    messagebox.showerror("Erro", "Coluna de processo não encontrada no backup.", parent=root)
    exit()

# Identifica coluna de status automaticamente
col_status = next((col for col in df_backup.columns if "status" in col.lower()), None)

# Extrai os processos já feitos
processos_feitos = df_backup[col_processo].dropna().astype(str).tolist()
total_feitos     = len(processos_feitos)

# Conta sucessos e falhas
processos_com_sucesso = 0
processos_sem_sucesso = 0
if col_status:
    processos_com_sucesso = int((df_backup[col_status] == "Mapeado").sum())
    processos_sem_sucesso = total_feitos - processos_com_sucesso

print(f"\n[OK] {total_feitos} processos encontrados no backup.")
print(f"     -> Mapeados com sucesso : {processos_com_sucesso}")
print(f"     -> Não mapeados / falhas: {processos_sem_sucesso}")

# --- Salvar checkpoint ---
checkpoint = {
    "arquivo_entrada": "",      # deixado vazio — obs6 vai pular a verificação de arquivo
    "processos_feitos": processos_feitos,
    "processos_com_sucesso": processos_com_sucesso,
    "processos_sem_sucesso": processos_sem_sucesso
}

with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
    json.dump(checkpoint, f, ensure_ascii=False)

messagebox.showinfo(
    "Checkpoint criado!",
    f"Checkpoint criado com sucesso!\n\n"
    f"{total_feitos} processos registrados.\n\n"
    f"Agora rode o obs6.py, selecione o arquivo de entrada\n"
    f"e clique 'Sim' para continuar de onde parou.",
    parent=root
)

print(f"\n[CONCLUÍDO] Checkpoint salvo em: data/{os.path.basename(CHECKPOINT_FILE)}")
print("Agora rode o obs6.py e clique 'Sim' para continuar. :)")
