from datetime import datetime
from config import LOG_FILE, SUCCESS_LOG


def _escrever(path, mensagem: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {mensagem}\n")


def log_erro(cnpj: str, motivo: str) -> None:
    _escrever(LOG_FILE, f"ERRO | {cnpj} | {motivo}")


def log_sucesso(cnpj: str) -> None:
    _escrever(SUCCESS_LOG, f"OK   | {cnpj}")
