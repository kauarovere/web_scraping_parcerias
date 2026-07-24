# 🔍 Observatório de Parcerias — Diário Oficial de SP

> Extrator automatizado de dados de **processos de parceria** publicados no [Diário Oficial do Município de São Paulo](https://diariooficial.prefeitura.sp.gov.br). A partir de uma lista de números de processo, o robô localiza e extrai: órgão, entidade parceira, objeto, modalidade, datas de vigência, dotação orçamentária e link do documento.

---

## 📋 Visão Geral

Dado um arquivo Excel/CSV com números de processo SEI, o sistema:
1. Busca cada processo no Diário Oficial
2. Localiza o documento correto (extrato, aditamento, prorrogação etc.)
3. Extrai os campos estruturados via parsing HTML
4. Gera uma planilha Excel consolidada com todos os dados

```
Arquivo de entrada (xlsx/csv)
  │  coluna: codProcesso_y
  ▼
  Formata número → 1234.5678/9012345-6
  │
  ▼
  POST no D.O. SP (busca por processo)
  │
  ▼
  Localiza documento relevante
  (extrato / retificação / aditamento / prorrogação...)
  │
  ▼
  GET no documento → parsing HTML
  │
  ▼
  Extrai: Órgão, Contratado, Contrato, Objeto,
          Modalidade, Datas, Dotação, Natureza
  │
  ▼
  data/Auditoria_obs6_Parcerias.xlsx
```

---

## 📁 Estrutura do Projeto

```
observatorio-parcerias-sp/
│
├── README.md
├── requirements.txt
├── .gitignore
│
├── src/                              ← Código-fonte
│   ├── obs6.py                       ← ⭐ Script principal (Turbo: paralelo + checkpoint)
│   ├── obs6_retry.py                 ← Reprocessador de falhas de conexão
│   └── obs6_criar_checkpoint.py      ← Reconstrói checkpoint a partir de backup
│
├── scripts/                          ← Atalhos Windows (.bat)
│   ├── EXECUTAR_OBS6.bat
│   ├── REPROCESSAR_FALHAS.bat
│   └── CRIAR_CHECKPOINT.bat
│
└── data/                             ← Dados gerados (ignorados pelo git)
    ├── Auditoria_obs6_Parcerias.xlsx ← Planilha final
    ├── Backup_Temp_obs6.xlsx         ← Backup intermediário automático
    └── obs6_checkpoint.json          ← Ponto de retomada
```

---

## 🚀 Como Usar

### 1. Instalação

```bash
git clone https://github.com/seu-usuario/observatorio-parcerias-sp.git
cd observatorio-parcerias-sp
pip install -r requirements.txt
```

### 2. Preparar o arquivo de entrada

Monte uma planilha `.xlsx` ou `.csv` com uma coluna chamada **`codProcesso_y`** contendo os números de processo (16 dígitos, com ou sem formatação):

| codProcesso_y |
|---|
| 6025202500269639 |
| 6025202500179079 |

### 3. Executar o obs6

```bash
python src/obs6.py
```

Uma janela abrirá para selecionar o arquivo de entrada. O progresso é exibido no terminal.

**Ou clique duas vezes em:** `scripts/EXECUTAR_OBS6.bat`

### 4. Retomar após interrupção

Se o processo foi interrompido (Ctrl+C ou queda de energia), simplesmente rode novamente:

```bash
python src/obs6.py
```

Selecione o **mesmo arquivo de entrada** e clique **"Sim"** para continuar.

### 5. Reprocessar falhas de conexão

```bash
python src/obs6_retry.py
```

Selecione a planilha de saída — o script refaz apenas os processos com erro.

### 6. Reconstruir checkpoint perdido

Se o checkpoint foi deletado mas o backup ainda existe:

```bash
python src/obs6_criar_checkpoint.py
```

---

## 📊 Colunas da Planilha de Saída

| Coluna | Descrição |
|---|---|
| `Nr do Processo` | Número formatado (ex: 6025.2025/0026963-9) |
| `Orgao` | Secretaria / órgão público responsável |
| `Nome do Contratado` | Entidade parceira (OSC) |
| `Nr do Contrato` | Número do instrumento |
| `Objeto do Contrato` | Descrição do objeto da parceria |
| `Modalidade` | Tipo (Termo de Fomento, Colaboração etc.) |
| `Data Inicio` | Início de vigência |
| `Data Fim` | Fim de vigência |
| `Dotacao Orcamentaria` | Dotação orçamentária |
| `Natureza da Despesa` | Natureza da despesa |
| `Status` | Mapeado / Não mapeado / Falha |
| `Link do Documento` | URL direta para o documento no D.O. |

---

## ⚙️ Configurações (topo do `obs6.py`)

```python
INTERVALO_BACKUP = 50    # salva backup a cada N processos
MAX_TENTATIVAS   = 3     # retries por processo em caso de falha
ESPERA_RETRY     = 20    # segundos de espera entre tentativas
NUM_WORKERS      = 2     # threads paralelas (↑ aumenta velocidade, ↑ carga no servidor)
SLEEP_ENTRE_REQ  = 0.5  # pausa entre requisições por thread (segundos)
```

---

## 🔧 Requisitos

- Python 3.9+
- Tkinter (incluso no Python padrão no Windows)

```bash
pip install -r requirements.txt
```

---

## 📄 Licença

Uso interno / fins de transparência pública. Os dados consultados são públicos.
