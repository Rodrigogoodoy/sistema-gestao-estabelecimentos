from pathlib import Path

# --- Arquivos de entrada ---
EXCEL_FILE = "050526__Base_Pesquisa_DataSUS.xlsx"
EXCEL_SHEET = 0
EXCEL_COLUMN = "cnpj"

# --- Site alvo ---
TARGET_URL = "https://cnes.datasus.gov.br/pages/estabelecimentos/consulta.jsp"

# --- Automação ---
HEADLESS = True                     # True = sem janela (mais rápido)
DELAY_BETWEEN_ENTRIES = 0.3        # segundos entre CNPJs encontrados
DELAY_ON_ERROR = 3.0               # segundos de espera após erro
MAX_RETRIES = 3                    # tentativas por CNPJ em caso de falha

# --- Arquivos de controle e saída ---
CHECKPOINT_FILE = Path("checkpoint.json")
LOG_FILE = Path("log_erros.txt")
SUCCESS_LOG = Path("log_sucesso.txt")
OUTPUT_FILE = Path("resultados.xlsx")
SAVE_EVERY = 100                   # salva o Excel a cada N registros
