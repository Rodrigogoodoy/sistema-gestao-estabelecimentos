import json
from config import CHECKPOINT_FILE


def salvar(indice: int) -> None:
    CHECKPOINT_FILE.write_text(json.dumps({"ultimo_indice": indice}), encoding="utf-8")


def carregar() -> int:
    if not CHECKPOINT_FILE.exists():
        return 0
    data = json.loads(CHECKPOINT_FILE.read_text(encoding="utf-8"))
    indice = data.get("ultimo_indice", 0)
    print(f"[checkpoint] Retomando do índice {indice}")
    return indice


def resetar() -> None:
    if CHECKPOINT_FILE.exists():
        CHECKPOINT_FILE.unlink()
    print("[checkpoint] Progresso resetado")
